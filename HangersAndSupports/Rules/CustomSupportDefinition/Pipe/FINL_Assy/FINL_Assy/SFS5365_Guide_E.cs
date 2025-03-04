//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5365_Guide_E.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5365_Guide_E
//   Author       :  BS
//   Creation Date:  08/07/2013 
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  08/07/2013    BS  CR-CP-224485- Converted HS_FINL_Assy to C# .Net
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Support.Middle;
using System.Linq;

namespace Ingr.SP3D.Content.Support.Rules
{
    public class SFS5365_Guide_E : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        CustomSupportDefinition Left, Bottom, Right,Top;
        string slideType;

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    slideType = (string)((PropertyValueString)support.GetPropertyValue("IJUAFINLSlideType", "SlideType")).PropValue;

                    Top = FINLAssemblyServices.GetAssembly("FINL_Assy,Ingr.SP3D.Content.Support.Rules." + slideType.Trim(), support);
                    Left = FINLAssemblyServices.GetAssembly("FINL_Assy,Ingr.SP3D.Content.Support.Rules." + slideType.Trim(), support);
                    Right = FINLAssemblyServices.GetAssembly("FINL_Assy,Ingr.SP3D.Content.Support.Rules." + slideType.Trim(), support);
                    Bottom = FINLAssemblyServices.GetAssembly("FINL_Assy,Ingr.SP3D.Content.Support.Rules." + slideType.Trim(), support);

                    FINLAssemblyServices.SetValueOnPropertyType(Bottom, "Index", 1);

                    FINLAssemblyServices.SetValueOnPropertyType(Top, "Override", true);
                    FINLAssemblyServices.SetValueOnPropertyType(Top, "Clamps", 2);
                    FINLAssemblyServices.SetValueOnPropertyType(Top, "Index", 2);

                    FINLAssemblyServices.SetValueOnPropertyType(Left, "Override", true);
                    FINLAssemblyServices.SetValueOnPropertyType(Left, "Clamps", 2);
                    FINLAssemblyServices.SetValueOnPropertyType(Left, "Index", 3);

                    FINLAssemblyServices.SetValueOnPropertyType(Right, "Override", true);
                    FINLAssemblyServices.SetValueOnPropertyType(Right, "Clamps", 2);
                    FINLAssemblyServices.SetValueOnPropertyType(Right, "Index", 4);

                    return new Collection<PartInfo>(Bottom.Parts.Concat(Top.Parts).Concat(Left.Parts).Concat(Right.Parts).ToList());

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
                    routeConnections = Bottom.SupportedConnections;
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
                    structConnections = Bottom.SupportingConnections;
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
                ReadOnlyCollection<object> topJoints = FINLAssemblyServices.GetSubAssemblyJoints(this, Top, oSupCompColl);
                ReadOnlyCollection<object> leftJoints = FINLAssemblyServices.GetSubAssemblyJoints(this, Left, oSupCompColl);
                ReadOnlyCollection<object> rightJoints = FINLAssemblyServices.GetSubAssemblyJoints(this, Right, oSupCompColl);
                ReadOnlyCollection<object> bottomJoints = FINLAssemblyServices.GetSubAssemblyJoints(this, Bottom, oSupCompColl);

                JointHelper.m_oCollOfJoints = new Collection<object>(topJoints.Concat(leftJoints).Concat(rightJoints).Concat(bottomJoints).ToList());

                Collection<ConnectionInfo> iSubAssyRouteConTop = Top.SupportedConnections;
                Collection<ConnectionInfo> iSubAssyRouteConLeft = Left.SupportedConnections;
                Collection<ConnectionInfo> iSubAssyRouteConRight = Right.SupportedConnections;

                JointHelper.CreateRigidJoint("-1", "Route", iSubAssyRouteConTop[0].PartKey, "Connection", Plane.YZ, Plane.YZ, Axis.Z, Axis.NegativeZ, 0, 0, 0);
                JointHelper.CreateRigidJoint("-1", "Route", iSubAssyRouteConLeft[0].PartKey, "Connection", Plane.YZ, Plane.YZ, Axis.Y, Axis.Z, 0, 0, 0);
                JointHelper.CreateRigidJoint("-1", "Route", iSubAssyRouteConRight[0].PartKey, "Connection", Plane.YZ, Plane.YZ, Axis.Y, Axis.NegativeZ, 0, 0, 0);

            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of SFS5365_Guide_E." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
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
                return BOMString = "Guide E SFS 5365 DN " + pipeInfo.NominalDiameter.Size;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - SFS5365_Guide_E" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion


    }

}