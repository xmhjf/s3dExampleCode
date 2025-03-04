//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5393D.cs
//    FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5393D
//   Author       :  Vijay
//   Creation Date:  18.03.2013
//   Description: CR-CP-222272 Convert HS_FINL_Parts VB Project to C# .Net  

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   18.03.2013     Vijay   CR-CP-222272 Convert HS_FINL_Parts VB Project to C# .Net 
//  11.Dec.2014      PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class SFS5393D : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5393D"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputString(2, "LoadClass", "LoadClass", "No Value")]
        public InputString m_oLoadClass;
        [InputDouble(3, "C", "C", 0.999999)]
        public InputDouble m_dC;
        [InputDouble(4, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(5, "H", "H", 0.999999)]
        public InputDouble m_dH;
        [InputDouble(6, "S", "S", 0.999999)]
        public InputDouble m_dS;
        [InputDouble(7, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(8, "D2", "D2", 0.999999)]
        public InputDouble m_dD2;
        [InputDouble(9, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(10, "W", "W", 0.999999)]
        public InputDouble m_dW;
        [InputDouble(11, "T", "T", 0.999999)]
        public InputDouble m_dT;
        [InputDouble(12, "L", "L", 0.999999)]
        public InputDouble m_dL;
        [InputDouble(13, "M", "M", 0.999999)]
        public InputDouble m_dM;
        [InputDouble(14, "F", "F", 0.999999)]
        public InputDouble m_dF;
        [InputDouble(15, "WorkingTrav", "WorkingTrav", 0.999999)]
        public InputDouble m_dWorkingTrav;
        [InputDouble(16, "Blocked", "Blocked", 0.999999)]
        public InputDouble m_oBlocked;
        [InputDouble(17, "Coefficient", "Coefficient", 0.999999)]
        public InputDouble m_dCoefficient;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("HOLE", "HOLE")]
        [SymbolOutput("LUG", "LUG")]
        [SymbolOutput("LUG2", "LUG2")]
        [SymbolOutput("JOINT", "JOINT")]
        public AspectDefinition m_Symbolic;

        #endregion

        #region "Construct Outputs"

        /// <summary>
        /// Construct symbol outputs in aspects.
        /// </summary>
        /// <remarks></remarks>
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;



                Double C = m_dC.Value;
                Double B = m_dB.Value;
                Double H = m_dH.Value;
                Double S = m_dS.Value;
                Double D = m_dD.Value;
                Double D2 = m_dD2.Value;
                Double A = m_dA.Value;
                Double W = m_dW.Value;
                Double T = m_dT.Value;
                Double L = m_dL.Value;
                Double M = m_dM.Value;
                Double F = m_dF.Value;
                Double workingTrav = m_dWorkingTrav.Value;
                Double coefficient = m_dCoefficient.Value;

                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Bottom", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Top", new Position(0, 0, L + F), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (C <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidCGTZero, "C value should be greater than zero"));
                    return;
                }
                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidDGTZero, "D value should be greater than zero"));
                    return;
                }
                if (T <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidTGTZero, "T value should be greater than zero"));
                    return;
                }
                if (B <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidBGTZero, "B value should be greater than zero"));
                    return;
                }
                if (W <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidWGTZero, "W value should be greater than zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(H + 2 * S, 0) == true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidHandS, "H + 2*S value cannot be zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(M + F, 0) == true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidMandF, "M + F value cannot be zero"));
                    return;
                }
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                //=================================================
                //Construction of Physical Aspect 
                //=================================================


                Vector normal = new Vector(0, 0, 1);
                symbolGeometryHelper.ActivePosition = new Position(0, 0, -A + W);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d bodyCylinder = symbolGeometryHelper.CreateCylinder(null, C / 2.0, H + S + S);
                m_Symbolic.Outputs["BODY"] = bodyCylinder;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-T / 2.0, -B / 2.0, -A);
                Projection3d lugBox = symbolGeometryHelper.CreateBox(null, T, B, W, 9);
                m_Symbolic.Outputs["LUG"] = lugBox;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(T / 2.0, 0, 0).Subtract(new Position(-T / 2.0, 0, 0));
                symbolGeometryHelper.ActivePosition = new Position(T / 2.0, 0, 0);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d holeCylinder = symbolGeometryHelper.CreateCylinder(null, D / 2.0, normal.Length);
                m_Symbolic.Outputs["HOLE"] = holeCylinder;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-B / 2.0, -T / 2.0, -A + W + H + S + S);
                Projection3d lug2Box = symbolGeometryHelper.CreateBox(null, B, T, W, 9);
                m_Symbolic.Outputs["LUG2"] = lug2Box;

                normal = new Vector(0, 0, 1);
                symbolGeometryHelper.ActivePosition = new Position(0, 0, L);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d jointCylinder = symbolGeometryHelper.CreateCylinder(null, D2 / 2.0, M + F);
                m_Symbolic.Outputs["JOINT"] = jointCylinder;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SFS5393D.cs."));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomString = "";
            try
            {
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                string material = (string)((PropertyValueString)part.GetPropertyValue("IJOAFINL_Material", "Material")).PropValue;
                double workingTrav = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJHgrPartWorkingTrav", "WorkingTrav")).PropValue;
                string loadClass = (string)((PropertyValueString)part.GetPropertyValue("IJUAFINL_LoadClass", "LoadClass")).PropValue;
                double H = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAFINL_H", "H")).PropValue;
                int blocked = (int)((PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJHgrPartBlocked", "Blocked")).PropValue;
                string strBlock;
                if (blocked == 1)
                    strBlock = "LP";
                else
                    strBlock = "";
                bomString = "Hanger spring SFS 5393 D " + loadClass + " " + Math.Round(H * 1000, 0) + " " + Math.Round(workingTrav * 1000, 0) + " " + strBlock;
                return bomString;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of SFS5393D"));
                }
                return "";
            }
        }

        #endregion
    }
}
