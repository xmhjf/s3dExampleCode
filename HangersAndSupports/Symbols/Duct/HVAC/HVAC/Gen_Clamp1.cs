//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Gen_Clamp1.cs
//    HVAC,Ingr.SP3D.Content.Support.Symbols.Gen_Clamp1
//   Author       :  Hema
//   Creation Date:  29-11-2012
//   Description:    Converted HS_HVAC VB Project to C# .Net

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   29-11-2012     Hema    CR-CP-222294 Converted HS_HVAC VB Project to C# .Net
//	 27/03/2013		Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   03/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
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

    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class Gen_Clamp1 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HVAC,Ingr.SP3D.Content.Support.Symbols.Gen_Clamp1"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Length", "Length", 0.999999)]
        public InputDouble m_dLength;
        [InputDouble(3, "ClampLength", "ClampLength", 0.999999)]
        public InputDouble m_dClampLength;
        [InputDouble(4, "LegLength", "LegLength", 0.999999)]
        public InputDouble m_dLegLength;
        [InputDouble(5, "Radius", "Radius", 0.999999)]
        public InputDouble m_dRadius;
        [InputDouble(6, "Thickness", "Thickness", 0.999999)]
        public InputDouble m_dThickness;
        [InputDouble(7, "Width", "Width", 0.999999)]
        public InputDouble m_dWidth;
        [InputDouble(8, "BoltDia", "BoltDia", 0.999999)]
        public InputDouble m_dBoltDia;
        [InputString(9, "InputBomDesc1", "InputBomDesc1", "No Value")]
        public InputString m_oInputBomDesc1;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("DuctClamp", "DuctClamp")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("Port4", "Port4")]
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
              
                double length = m_dLength.Value;
                double legLength = m_dLegLength.Value;
                double radius = m_dRadius.Value;
                double thickness = m_dThickness.Value;
                double width = m_dWidth.Value;
                double A = 2 * (radius + legLength);
                double E = radius + thickness + length;

                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;
                Port port3 = new Port(OccurrenceConnection, part, "StartOther", new Position(0, radius + legLength / 2, length - thickness / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;
                Port port4 = new Port(OccurrenceConnection, part, "EndOther", new Position(0, -radius - legLength / 2, length - thickness / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port4"] = port4;

                if (width == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HVACLocalizer.GetString(HVACSymbolResourceIDs.ErrInvalidWidth, "Width cannot be zero"));
                    return;
                }
                if (radius == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HVACLocalizer.GetString(HVACSymbolResourceIDs.ErrInvalidRadius, "Radius cannot be zero"));
                    return;
                }
                if (legLength == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HVACLocalizer.GetString(HVACSymbolResourceIDs.ErrInvalidlegLength, "Leg Length cannot be zero"));
                    return;
                }
                if (thickness <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HVACLocalizer.GetString(HVACSymbolResourceIDs.ErrInvalidthickness, "Thickness should be greater than zero"));
                    return;
                }
                //=================================================
                //Construction of Physical Aspect 
                //=================================================                

                Collection<ICurve> curveCollection = new Collection<ICurve>();
                Matrix4X4 matrix = new Matrix4X4();

                curveCollection.Add(new Line3d( new Position(-width / 2, length - E, 0), new Position(-width / 2, length - E, length - thickness)));
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                matrix.SetIdentity();
                matrix.Rotate(Math.PI, new Vector(0, 0, 1));

                Arc3d outerArc = symbolGeometryHelper.CreateArc(null, E - length, Math.PI);
                outerArc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                outerArc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Translate(new Vector(0, -width / 2, 0));
                outerArc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                outerArc.Transform(matrix);
                curveCollection.Add(outerArc);

                curveCollection.Add(new Line3d( new Position(-width / 2, E - length, 0), new Position(-width / 2, E - length, length - thickness)));
                curveCollection.Add(new Line3d( new Position(-width / 2, A / 2, length - thickness), new Position(-width / 2, E - length, length - thickness)));
                curveCollection.Add(new Line3d( new Position(-width / 2, A / 2, length), new Position(-width / 2, A / 2, length - thickness)));
                curveCollection.Add(new Line3d( new Position(-width / 2, E - length - thickness, length), new Position(-width / 2, A / 2, length)));
                curveCollection.Add(new Line3d( new Position(-width / 2, E - length - thickness, 0), new Position(-width / 2, E - length - thickness, length)));

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                matrix.SetIdentity();
                matrix.Rotate(Math.PI, new Vector(0, 0, 1));

                Arc3d innerArc = symbolGeometryHelper.CreateArc(null, E - length - thickness, Math.PI);
                innerArc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                innerArc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Translate(new Vector(0, -width / 2, 0));
                innerArc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                innerArc.Transform(matrix);
                curveCollection.Add(innerArc);

                curveCollection.Add(new Line3d( new Position(-width / 2, length - E + thickness, 0), new Position(-width / 2, length - E + thickness, length)));
                curveCollection.Add(new Line3d( new Position(-width / 2, length - E + thickness, length), new Position(-width / 2, -A / 2, length)));
                curveCollection.Add(new Line3d( new Position(-width / 2, -A / 2, length), new Position(-width / 2, -A / 2, length - thickness)));
                curveCollection.Add(new Line3d( new Position(-width / 2, -A / 2, length - thickness), new Position(-width / 2, length - E, length - thickness)));

                Projection3d ductClamp = new Projection3d(new ComplexString3d(curveCollection),  new Vector(1, 0, 0), width, true);
                m_Symbolic.Outputs["DuctClamp"] = ductClamp;
            }
            catch    //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HVACLocalizer.GetString(HVACSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Gen_Clamp1"));
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

                string bomDescription = (String)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrHVACBomDesc", "InputBomDesc1")).PropValue;

                if (bomDescription.Trim() == "None")
                    bomString = " ";
                else if (bomDescription.Trim() == " ")
                    bomString = part.PartDescription;

                return bomString;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, HVACLocalizer.GetString(HVACSymbolResourceIDs.ErrConstructOutputs, "Error in BOMDescription of Gen_Clamp1"));
                }
                return "";
            }
        }
        #endregion
    }
}
