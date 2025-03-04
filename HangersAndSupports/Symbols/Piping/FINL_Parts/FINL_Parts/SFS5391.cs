//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5391.cs
//   FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5391
//   Author       :  Hema
//   Creation Date:  18-03-2013
//   Description:    Converted FINL_Parts VB Project to C#.Net Project 

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   18-03-2013      Hema    Converted FINL_Parts VB Project to C#.Net Project 
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
    [SymbolVersion("1.0.0.0")]
    [CacheOption(CacheOptionType.Cached)]
    public class SFS5391 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5391"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputString(2, "LoadClass", "LoadClass", "No Value")]
        public InputString m_oLoadClass;
        [InputDouble(3, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(4, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(5, "C", "C", 0.999999)]
        public InputDouble m_dC;
        [InputDouble(6, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(7, "F", "F", 0.999999)]
        public InputDouble m_dF;
        [InputDouble(8, "G", "G", 0.999999)]
        public InputDouble m_dG;
        [InputDouble(9, "H", "H", 0.999999)]
        public InputDouble m_dH;
        [InputDouble(10, "S", "S", 0.999999)]
        public InputDouble m_dS;
        [InputDouble(11, "X", "X", 0.999999)]
        public InputDouble m_dX;
        [InputDouble(12, "L", "L", 0.999999)]
        public InputDouble m_dL;
        [InputDouble(13, "K1", "K1", 0.999999)]
        public InputDouble m_dK1;
        [InputDouble(14, "K2", "K2", 0.999999)]
        public InputDouble m_dK2;
        [InputDouble(15, "RodDia", "RodDia", 0.999999)]
        public InputDouble m_dRodDia;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("TOP", "TOP")]
        [SymbolOutput("BOT", "BOT")]
        [SymbolOutput("ROD", "ROD")]
        [SymbolOutput("LEFT", "LEFT")]
        [SymbolOutput("RIGHT", "RIGHT")]
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



                Double A = m_dA.Value;
                Double B = m_dB.Value;
                Double C = m_dC.Value;
                Double D = m_dD.Value;
                Double F = m_dF.Value;
                Double G = m_dG.Value;
                Double H = m_dH.Value;
                Double S = m_dS.Value;
                Double X = m_dX.Value;
                Double L = m_dL.Value;
                Double K1 = m_dK1.Value;
                Double K2 = m_dK2.Value;
                Double rodDiameter = m_dRodDia.Value;

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "BotExThdRH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, 0, L - X + A - X), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidDGTZero, "D value should be greater than zero"));
                    return;
                }
                if (S <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidSGTZero, "S value should be greater than zero"));
                    return;
                }
                if (rodDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidRodDiameterGTZero, "Rod Diameter should be greater than zero"));
                    return;
                }
                //=================================================
                //Construction of Physical Aspect 
                //=================================================

                //Create the Graphics

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(0, 0, L).Subtract(new Position(0, 0, 0));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d rod = (Projection3d)symbolGeometryHelper.CreateCylinder(null, rodDiameter / 2.0, normal.Length);
                m_Symbolic.Outputs["ROD"] = rod;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal1 = new Position(0, 0, L - X + H).Subtract(new Position(0, 0, L - X));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, L - X);
                symbolGeometryHelper.SetOrientation(normal1, normal1.GetOrthogonalVector());
                Projection3d bottom = (Projection3d)symbolGeometryHelper.CreateCylinder(null, D / 2.0, normal1.Length);
                m_Symbolic.Outputs["BOT"] = bottom;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal3 = new Position(0, 0, L - X + A).Subtract(new Position(0, 0, L - X + A - H));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, L - X + A - H);
                symbolGeometryHelper.SetOrientation(normal3, normal3.GetOrthogonalVector());
                Projection3d top = (Projection3d)symbolGeometryHelper.CreateCylinder(null, D / 2.0, normal3.Length);
                m_Symbolic.Outputs["TOP"] = top;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-C / 2.0, -S / 2.0, L - X + F);
                Projection3d left = (Projection3d)symbolGeometryHelper.CreateBox(null, (C - G) / 2.0, S, A - (F * 2.0), 9);
                m_Symbolic.Outputs["LEFT"] = left;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(G / 2.0, -S / 2.0, L - X + F);
                Projection3d right = (Projection3d)symbolGeometryHelper.CreateBox(null, (C - G) / 2.0, S, A - (F * 2.0), 9);
                m_Symbolic.Outputs["RIGHT"] = right;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SFS5391."));
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
                String loadNumber = (String)((PropertyValueString)part.GetPropertyValue("IJUAFINL_LoadClass", "LoadClass")).PropValue;

                bomString = "Turnbuckle K SFS 5391- " + loadNumber;

                return bomString;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of SFS5391"));
                }
                return "";
            }

        }
        #endregion
    }
}
