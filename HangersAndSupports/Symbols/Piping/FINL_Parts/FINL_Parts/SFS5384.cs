//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5384.cs
//   FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5384
//   Author       :  Hema
//   Creation Date:  18-03-2013
//   Description:    Converted FINL_Parts VB Project to C#.Net Project 

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   18-03-2013      Hema    Converted FINL_Parts VB Project to C#.Net Project 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
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
    public class SFS5384 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5384"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(3, "S", "S", 0.999999)]
        public InputDouble m_dS;
        [InputDouble(4, "C", "C", 0.999999)]
        public InputDouble m_dC;
        [InputDouble(5, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(6, "E", "E", 0.999999)]
        public InputDouble m_dE;
        [InputDouble(7, "F", "F", 0.999999)]
        public InputDouble m_dF;
        [InputDouble(8, "G", "G", 0.999999)]
        public InputDouble m_dG;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("Port4", "Port4")]
        [SymbolOutput("Port5", "Port5")]
        [SymbolOutput("HOLE1", "HOLE1")]
        [SymbolOutput("HOLE2", "HOLE2")]
        [SymbolOutput("HOLE3", "HOLE3")]
        [SymbolOutput("HOLE22", "HOLE22")]
        [SymbolOutput("HOLE33", "HOLE33")]
        [SymbolOutput("BODY", "BODY")]
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
                Double S = m_dS.Value;
                Double C = m_dC.Value;
                Double D = m_dD.Value;
                Double E = m_dE.Value;
                Double F = m_dF.Value;
                Double G = m_dG.Value;

                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Hole1", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Hole2", new Position(0, 0, -C), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;
                Port port3 = new Port(OccurrenceConnection, part, "Hole3", new Position(-C, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;
                Port port4 = new Port(OccurrenceConnection, part, "Hole22", new Position(0, 0, -C), new Vector(Math.Sqrt(2) / 2, 0, -Math.Sqrt(2) / 2), new Vector(Math.Sqrt(2) / 2, 0, Math.Sqrt(2) / 2));
                m_Symbolic.Outputs["Port4"] = port4;
                Port port5 = new Port(OccurrenceConnection, part, "Hole33", new Position(-C, 0, 0), new Vector(Math.Sqrt(2) / 2, 0, -Math.Sqrt(2) / 2), new Vector(Math.Sqrt(2) / 2, 0, Math.Sqrt(2) / 2));
                m_Symbolic.Outputs["Port5"] = port5;

                if (S == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidSNEZ, "S value cannot be zero"));
                    return;
                }
                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidDGTZero, "D value should be greater than zero"));
                    return;
                }

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
               
                //Create the Graphics
                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(E - G, -S / 2.0, E));
                pointCollection.Add(new Position(-C - E, -S / 2.0, E));
                pointCollection.Add(new Position(-C - E, -S / 2.0, E - F));
                pointCollection.Add(new Position(E - F, -S / 2.0, -C - E));
                pointCollection.Add(new Position(E, -S / 2.0, -C - E));
                pointCollection.Add(new Position(E, -S / 2.0, E - G));
                pointCollection.Add(new Position(E - G, -S / 2.0, E));
              
                Projection3d body = new Projection3d(new LineString3d(pointCollection),new Vector(0, 1, 0),S, true);
                m_Symbolic.Outputs["BODY"] = body;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(0, S / 2.0, 0).Subtract(new Position(0, -S / 2.0, 0));
                symbolGeometryHelper.ActivePosition = new Position(0, -S / 2.0, 0);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d hole1 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, D / 2.0, normal.Length);
                m_Symbolic.Outputs["HOLE1"] = hole1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal1 = new Position(-C, S / 2.0, 0).Subtract(new Position(-C, -S / 2.0, 0));
                symbolGeometryHelper.ActivePosition = new Position(-C, -S / 2.0, 0);
                symbolGeometryHelper.SetOrientation(normal1, normal1.GetOrthogonalVector());
                Projection3d hole2 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, D / 2.0, normal1.Length);
                m_Symbolic.Outputs["HOLE2"] = hole2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal2 = new Position(0, S / 2.0, -C).Subtract(new Position(0, -S / 2.0, -C));
                symbolGeometryHelper.ActivePosition = new Position(0, -S / 2.0, -C);
                symbolGeometryHelper.SetOrientation(normal2, normal2.GetOrthogonalVector());
                Projection3d hole3 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, D / 2.0, normal2.Length);
                m_Symbolic.Outputs["HOLE3"] = hole3;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SFS5384."));
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

                bomString = "Triangle plate SFS 5384 - " + loadNumber;

                return bomString;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of SFS5384"));
                }
                return "";
            }
        }
        #endregion
    }
}
