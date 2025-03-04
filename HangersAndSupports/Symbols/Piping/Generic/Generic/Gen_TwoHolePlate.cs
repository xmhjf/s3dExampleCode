//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Gen_TwoHolePlate.cs
//    Generic,Ingr.SP3D.Content.Support.Symbols.Gen_TwoHolePlate
//   Author       :  Hema
//   Creation Date:  14-11-2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   14-11-2012      Hema    CR-CP-222274  Converted HS_Generic VB Project to C# .Net 
//	 27/03/2013		Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   6/11/2013     Vijaya   CR-CP-242533  Provide the ability to store GType outputs as a single blob for H&S .net symbols  
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
    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class Gen_TwoHolePlate : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Generic,Ingr.SP3D.Content.Support.Symbols.Gen_TwoHolePlate"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "H1", "H1", 0.999999)]
        public InputDouble m_dH1;
        [InputDouble(3, "T1", "T1", 0.999999)]
        public InputDouble m_dT1;
        [InputDouble(4, "L", "L", 0.999999)]
        public InputDouble m_dL;
        [InputDouble(5, "H", "H", 0.999999)]
        public InputDouble m_dH;
        [InputDouble(6, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputString(7, "BOM_DESC1", "BOM_DESC1", "No Value")]
        public InputString m_dBOM_DESC1;
        [InputDouble(8, "L1", "L1", 0.999999)]
        public InputDouble m_dL1;
     
        #endregion

        #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("LINE", "LINE")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("R_HOLE", "R_HOLE")]
        [SymbolOutput("L_HOLE", "L_HOLE")]
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
               
                Double H1 = m_dH1.Value;
                Double T = m_dT1.Value;                
                Double L = m_dL.Value;                
                Double H = m_dH.Value;                
                Double D = m_dD.Value;                
                Double L1 = m_dL1.Value;                

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                
                Port port1 = new Port(OccurrenceConnection, part, "TopStructure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "BotStructure", new Position(0, 0, T), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (T == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrInvalidT1, "Plate Thickness cannot be zero"));
                    return;
                }
                if (L <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrInvalidLGTZero, "Plate Width should be greater than zero"));
                    return;
                }
                if (H <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrInvalidHGTZero, "Plate Depth should be greater than zero"));
                    return;
                }
                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrInvalidDGTZero, "Hole Diameter should be greater than zero"));
                    return;
                }

                symbolGeometryHelper.ActivePosition = new Position(0,0,0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 1, 0));
                Projection3d body = (Projection3d)symbolGeometryHelper.CreateBox(null, T, L ,H);
                m_Symbolic.Outputs["BODY"] = body;

                Line3d line = new Line3d(new Position(-H / 2.0, 0, T / 2.0), new Position(H / 2.0, 0, T / 2.0));
                m_Symbolic.Outputs["LINE"] = line;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(H / 2 - H1, L / 2 - L1, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d rightHole = symbolGeometryHelper.CreateCylinder(null, D / 2.0, T + 0.00002);
                m_Symbolic.Outputs["R_HOLE"] = rightHole;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(H / 2 - H1, -(L / 2 - L1), 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d leftHole = symbolGeometryHelper.CreateCylinder(null, D / 2.0, T + 0.00002);
                m_Symbolic.Outputs["L_HOLE"] = leftHole;
            }
            catch    //General Unhandled exception
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrConstructOutputs, "Error in ConstructOutputs of Gen_TwoHolePlate"));
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
                Part part = (Part)oSupportOrComponent.GetRelationship("madefrom", "part").TargetObjects[0];
                double LValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrGenTwoHolePlate", "L")).PropValue;
                double HValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrGenTwoHolePlate", "H")).PropValue;
                double DValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrGenTwoHolePlate", "D")).PropValue;
                double H1Value = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrGenTwoHolePlate", "H1")).PropValue;
                double TValue = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrGenTwoHolePlate", "T1")).PropValue;

                string bomDescription = (string)((PropertyValueString)part.GetPropertyValue("IJOAHgrGenPartBOMDesc", "BOM_DESC1")).PropValue;
                
                string L = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, LValue, UnitName.DISTANCE_MILLIMETER);
                string H = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, HValue, UnitName.DISTANCE_MILLIMETER);
                string D = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, DValue, UnitName.DISTANCE_MILLIMETER);
                string H1 = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, H1Value, UnitName.DISTANCE_MILLIMETER);
                string T = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, TValue, UnitName.DISTANCE_MILLIMETER);

                if (bomDescription == "None")
                    bomString = " ";
                else
                    bomString = T + " Plate Steel, " + L + " X " + H + " with Two " + D + " Holes with " + H1 + " inset";

                return bomString;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrBOMDescription, "Error in BOMDescription of Gen_TwoHolePlate"));
                }
                return "";
            }
        }
        #endregion

        #region "ICustomWeightCG Members"

        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                Double weight, cogX = 0, cogY = 0, cogZ = 0;
                const int getSteelDensityKGPerM = 7900;
                Part part = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];

                double width = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrGenTwoHolePlate", "L")).PropValue;
                double depth = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrGenTwoHolePlate", "H")).PropValue;
                double holeSize = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrGenTwoHolePlate", "D")).PropValue;
                double T = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrGenTwoHolePlate", "T1")).PropValue;

                weight = ((depth * width * T) - (2.0 * (holeSize / 2.0 * holeSize / 2.0 * T * Math.PI))) * getSteelDensityKGPerM;
               
                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrWeightCG, "Error in WeightCG of Gen_TwoHolePlate"));
                }
            }
        }
        #endregion
    }
}
