//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5388.cs
//   FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5388
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
    public class SFS5388 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5388"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Width", "Width", 0.999999)]
        public InputDouble m_dWidth;
        [InputDouble(3, "Thickness", "Thickness", 0.999999)]
        public InputDouble m_dThickness;
        [InputDouble(4, "HoleDia", "HoleDia", 0.999999)]
        public InputDouble m_dHoleDia;
        [InputDouble(5, "HoleCtoC", "HoleCtoC", 0.999999)]
        public InputDouble m_dHoleCtoC;
        [InputDouble(6, "HoleInset", "HoleInset", 0.999999)]
        public InputDouble m_dHoleInset;
        [InputDouble(7, "UserLength", "UserLength", 0.999999)]
        public InputDouble m_dUserLength;
        [InputDouble(8, "TakeOut", "TakeOut", 0.999999)]
        public InputDouble m_dTakeOut;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("VerPlate", "VerPlate")]
        [SymbolOutput("HorPlate", "HorPlate")]
        [SymbolOutput("Hole1", "Hole1")]
        [SymbolOutput("Hole2", "Hole2")]
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
                 
               
                
                Double width = m_dWidth.Value;
                Double thickness = m_dThickness.Value;
                Double holeDiameter = m_dHoleDia.Value;
                Double holeCtoC = m_dHoleCtoC.Value;
                Double holeInset = m_dHoleInset.Value;
                Double length = m_dUserLength.Value;
                Double takeOut = m_dTakeOut.Value;
                Matrix4X4 matrix = new Matrix4X4();

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Hole1", new Position(-(length / 2 - 0.05), holeInset, width - (takeOut + thickness)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Hole2", new Position(length / 2 - 0.05, holeInset, width - (takeOut + thickness)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;
                Port port3 = new Port(OccurrenceConnection, part, "Route", new Position(0, width / 2, width), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;


                if (length <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidLengthGTZero, "Length should be greater than zero"));
                    return;
                }
                if (width <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidWidthGTZero, "Width should be greater than zero"));
                    return;
                }
                if (holeDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidHoleDiameterGTZero, "Hole Dia value should be greater than zero"));
                    return;
                }
                if (thickness == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidThickness, "Thickness cannot be zero"));
                    return;
                }

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-length / 2, 0, -thickness);
                Projection3d verPlate = (Projection3d)symbolGeometryHelper.CreateBox(null, length, width - thickness, thickness, 9);
                matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                verPlate.Transform(matrix);
                m_Symbolic.Outputs["VerPlate"] = verPlate;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-length / 2, 0, width - thickness);
                Projection3d horPlate = (Projection3d)symbolGeometryHelper.CreateBox(null, length, width, thickness, 9);
                m_Symbolic.Outputs["HorPlate"] = horPlate;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-holeInset, -(length / 2 - 0.05), width - thickness);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), (new Vector(0, 0, 1)).GetOrthogonalVector());
                Projection3d hole1 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, holeDiameter / 2, thickness);
                matrix = new Matrix4X4();
                matrix.Rotate(3*Math.PI / 2, new Vector(0, 0, 1));
                hole1.Transform(matrix);
                m_Symbolic.Outputs["Hole1"] = hole1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-holeInset, (length / 2 - 0.05), width - thickness);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), (new Vector(0, 0, 1)).GetOrthogonalVector());
                Projection3d hole2 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, holeDiameter / 2, thickness);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                hole2.Transform(matrix);
                m_Symbolic.Outputs["Hole2"] = hole2;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SFS5388."));
                    return;
                }
            }
        }
        #endregion

        //In RefData BOM type is 2,In VB the implementation was in the same way 

        //#region "ICustomHgrBOMDescription Members"

        //public string BOMDescription(BusinessObject oSupportOrComponent)
        //{
        //    string bomString = "";
        //    try
        //    {
        //        double size, lengthValue;
        //        Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

        //        size = (double)((PropertyValueDouble)part.GetPropertyValue("", "SIZE")).PropValue;
        //        lengthValue = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrPSL_RC6", "F")).PropValue;
        //        String length = (MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lengthValue, UnitName.DISTANCE_MILLIMETER)).Trim();

        //        bomString = "PSL " + size + " Pressed Riser Clamp Six Bolt Type, Rod Centre =" + length;

        //        return bomString;
        //    }
        //    catch
        //    {
        //        if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
        //        {
        //            ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of SFS5388"));
        //        }
        //        return "";
        //    }
        //}
        //#endregion
    }
}
