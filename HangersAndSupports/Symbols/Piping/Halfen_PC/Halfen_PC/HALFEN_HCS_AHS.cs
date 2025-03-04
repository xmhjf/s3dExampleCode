//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   HALFEN_HCS_AHS.cs
//    Halfen_PC,Ingr.SP3D.Content.Support.Symbols.HALFEN_HCS_AHS
//   Author       : Sasidhar 
//   Creation Date: 16-11-2012 
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   16-11-2012     Sasidhar  CR-CP-222275 Converted VB HS_HALFEN_PC Project to C#.Net
//   20-03-2013      Vijay    DI-CP-228142  Modify the error handling for delivered H&S symbols  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;

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

    public class HALFEN_HCS_AHS : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Halfen_PC,Ingr.SP3D.Content.Support.Symbols.HALFEN_HCS_AHS"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Width", "Width", 0.999999)]
        public InputDouble m_dWidth;
        [InputDouble(3, "Depth", "Depth", 0.999999)]
        public InputDouble m_dDepth;
        [InputDouble(4, "Height", "Height", 0.999999)]
        public InputDouble m_dHeight;
        [InputDouble(5, "Bolt_Dia", "Bolt_Dia", 0.999999)]
        public InputDouble m_dBolt_Dia;
        [InputDouble(6, "Bolt_L", "Bolt_L", 0.999999)]
        public InputDouble m_dBolt_L;
        [InputDouble(7, "Bolt_Inset", "Bolt_Inset", 0.999999)]
        public InputDouble m_dBolt_Inset;
        [InputString(8, "Finish", "Finish", "No Value")]
        public InputString m_oFinish;
        [InputString(9, "Stock_Number", "Stock_Number", "No Value")]
        public InputString m_oStock_Number;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("CLIP", "CLIP")]
        [SymbolOutput("BOLT", "BOLT")]
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
                Part part =(Part) m_PartInput.Value;
                SP3DConnection connection = default(SP3DConnection);
                connection = OccurrenceConnection;

                Double width = m_dWidth.Value;
                Double depth = m_dDepth.Value;
                Double height = m_dHeight.Value;
                Double boltDiameter = m_dBolt_Dia.Value;
                Double boltL = m_dBolt_L.Value;
                Double boltInset = m_dBolt_Inset.Value;
                const Double width1 = 0.03;
                const Double height1 = 0.01;

                if (boltDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrBoltDiameter, "BoltDiameter can not be zero or negative"));
                    return;
                }
                if (boltL == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrBoltL, "BoltL can not be zero"));
                    return;
                }

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(connection, part, "Other", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Line3d line;
                Collection<ICurve> curvecollection = new Collection<ICurve>();

                line = new Line3d(new Position(-depth / 2.0, -boltInset, 0), new Position(-depth / 2.0, -boltInset + width1, 0));
                curvecollection.Add(line);

                line = new Line3d(new Position(-depth / 2.0, -boltInset + width1, 0), new Position(-depth / 2.0, -boltInset + width1, height1));
                curvecollection.Add(line);

                line = new Line3d(new Position(-depth / 2.0, -boltInset + width1, height1), new Position(-depth / 2.0, -boltInset + width, height1));
                curvecollection.Add(line);

                line = new Line3d(new Position(-depth / 2.0, -boltInset + width, height1), new Position(-depth / 2.0, -boltInset + width, height));
                curvecollection.Add(line);

                line = new Line3d(new Position(-depth / 2.0, -boltInset + width, height), new Position(-depth / 2.0, -boltInset, height));
                curvecollection.Add(line);

                line = new Line3d(new Position(-depth / 2.0, -boltInset, height), new Position(-depth / 2.0, -boltInset, 0));
                curvecollection.Add(line);

                Vector lineVector = new Vector(1, 0, 0);
                Projection3d body = new Projection3d(new ComplexString3d(curvecollection), lineVector, depth, true);
                m_Symbolic.Outputs["CLIP"] = body;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, -boltL + height + 3 * boltDiameter / 4.0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d cylinder = (Projection3d)symbolGeometryHelper.CreateCylinder(null, boltDiameter / 2.0, boltL);
                m_Symbolic.Outputs["BOLT"] = cylinder;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of HALFEN_HCS_AHS.cs."));
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
                String partSize;
                String finish;
                String stockNumber;
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                finish = (String)((PropertyValueString)part.GetPropertyValue("IJUAHgrFinish", "Finish")).PropValue;
                stockNumber = (String)((PropertyValueString)part.GetPropertyValue("IJUAHgrStock_Number", "Stock_Number")).PropValue;
                partSize = "5/1";
                bomString = part.PartDescription + " " + partSize + " - " + finish + " - " + stockNumber;
                return bomString;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrBOMDescription, "Error in BOMDescription of HALFEN_HCS_AHS.cs."));
                }
                return "";
            }
        }
        #endregion

    }

}
