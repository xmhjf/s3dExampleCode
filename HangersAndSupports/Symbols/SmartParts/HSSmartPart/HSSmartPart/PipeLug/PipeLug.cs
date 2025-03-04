//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   PipeLug.cs
//   HSSmartPart,Ingr.SP3D.Content.Support.Symbols.PipeLug
//   Author       :  Vinod Peddi
//   Creation Date:  30-Nov-2015
//   DI-CP-282644  Integrate the newly developed SmartParts into Product 
//   
//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
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
    [VariableOutputs]
    public class PipeLug : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.PipeLug"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;

        #endregion

        #region "Definition of Additional Inputs"

        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                List<Input> additionalInputs = new List<Input>();

                int startIndex = 2;
                additionalInputs.Add((Input)new InputDouble(startIndex, "RodDiameter", "RodDiameter", 0, false));
                additionalInputs.Add((Input)new InputDouble(++startIndex, "PipeOD", "PipeOD", 0, false));
                additionalInputs.Add((Input)new InputDouble(++startIndex, "RodTakeOut", "RodTakeOut", 0, false));
                additionalInputs.Add((Input)new InputDouble(++startIndex, "IsPerpendicular", "IsPerpendicular", 0, false));
                additionalInputs.Add((Input)new InputDouble(++startIndex, "TopShape", "TopShape", 1, false));
                additionalInputs.Add((Input)new InputDouble(++startIndex, "Angle1", "Angle1", 0, false));
                additionalInputs.Add((Input)new InputDouble(++startIndex, "Width1", "Width1", 0, false));
                additionalInputs.Add((Input)new InputDouble(++startIndex, "Width2", "Width2", 0, false));
                additionalInputs.Add((Input)new InputDouble(++startIndex, "Thickness1", "Thickness1", 0, false));
                additionalInputs.Add((Input)new InputDouble(++startIndex, "Angle2", "Angle2", 0, false));
                additionalInputs.Add((Input)new InputDouble(++startIndex, "Angle3", "Angle3", 0, false));
                additionalInputs.Add((Input)new InputDouble(++startIndex, "Gap1", "Gap1", 0, false));
                additionalInputs.Add((Input)new InputDouble(++startIndex, "StiffenerOffset", "StiffenerOffset", 0, false));
                additionalInputs.Add((Input)new InputDouble(++startIndex, "StiffenerHeight", "StiffenerHeight", 0, false));
                additionalInputs.Add((Input)new InputDouble(++startIndex, "StiffenerLength", "StiffenerLength", 0, false));
                additionalInputs.Add((Input)new InputDouble(++startIndex, "Offset1", "Offset1", 0, false));
                additionalInputs.Add((Input)new InputDouble(++startIndex, "ChamfLength", "ChamfLength", 0, false));
                additionalInputs.Add((Input)new InputDouble(++startIndex, "Pin1Diameter", "Pin1Diameter", 0, false));
                additionalInputs.Add((Input)new InputDouble(++startIndex, "Pin1Length", "Pin1Length", 0, false));
                additionalInputs.Add((Input)new InputDouble(++startIndex, "Height2", "Height2", 0, false));
                return additionalInputs;
            }
        }

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        //Ports as Outputs
        [SymbolOutput("Route", "Route")]
        [SymbolOutput("Lug", "Lug")]

        public AspectDefinition m_PhysicalAspect;

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
                //=================================================
                //Construction of Physical Aspect 
                //================================================= 

                //int outindex = new int;

                PipeLugInputs pipeLugInputs = LoadPipeLugInputsData(2);
                if (base.ToDoListMessage != null)
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        return;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                Matrix4X4 matrix = new Matrix4X4();

                if (pipeLugInputs.IsPerpendicular == false)
                {
                    //ports
                    Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_PhysicalAspect.Outputs["Route"] = port1;

                    Port port2 = new Port(OccurrenceConnection, part, "Lug", new Position(0, 0, pipeLugInputs.RodTakeOut), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_PhysicalAspect.Outputs["Lug"] = port2;


                    //Add Elbow geometry
                    if (HgrCompareDoubleService.cmpdbl(pipeLugInputs.Gap1,0) == true) //For Single Lug
                        AddWeldingLug(pipeLugInputs, matrix, m_PhysicalAspect.Outputs, "pipeLugShape1");
                    else  //For Double Lug
                    {
                        matrix.Translate(new Vector(0, pipeLugInputs.Gap1 / 2 + pipeLugInputs.Thickness1 / 2, 0));
                        AddWeldingLug(pipeLugInputs, matrix, m_PhysicalAspect.Outputs, "pipeLugShape1");
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Translate(new Vector(0, -pipeLugInputs.Gap1 / 2 - pipeLugInputs.Thickness1 / 2, 0));
                        AddWeldingLug(pipeLugInputs, matrix, m_PhysicalAspect.Outputs, "pipeLugShape2");

                        //Add Stiffener
                        if (HgrCompareDoubleService.cmpdbl(pipeLugInputs.StiffenerHeight, 0) == false && HgrCompareDoubleService.cmpdbl(pipeLugInputs.StiffenerLength, 0) == false)
                        {
                            symbolGeometryHelper.ActivePosition = new Position(-pipeLugInputs.StiffenerLength / 2, -pipeLugInputs.Gap1 / 2, pipeLugInputs.RodTakeOut - pipeLugInputs.StiffenerHeight - pipeLugInputs.StiffenerOffset);
                            BusinessObject stiffener = symbolGeometryHelper.CreateBox(OccurrenceConnection, pipeLugInputs.StiffenerLength, pipeLugInputs.Gap1, pipeLugInputs.StiffenerHeight, 9);
                            m_PhysicalAspect.Outputs["pipeLugStiffener"] = stiffener;
                        }
                        //Add Pin
                        if (HgrCompareDoubleService.cmpdbl(pipeLugInputs.Pin1Diameter, 0) == false && HgrCompareDoubleService.cmpdbl(pipeLugInputs.Pin1Length, 0) == false)
                        {
                            symbolGeometryHelper = new SymbolGeometryHelper();
                            symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                            symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Rotate((Math.PI / 2), new Vector(1, 0, 0));
                            Projection3d pipeLugPin = new Projection3d(symbolGeometryHelper.CreateCylinder(null, pipeLugInputs.Pin1Diameter / 2, pipeLugInputs.Pin1Length));
                            pipeLugPin.Transform(matrix);
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(0, pipeLugInputs.Pin1Length / 2,  pipeLugInputs.RodTakeOut));
                            pipeLugPin.Transform(matrix);
                            m_PhysicalAspect.Outputs["pipeLugPin"] = pipeLugPin;
                        }
                    }
                }


                if (pipeLugInputs.IsPerpendicular == true)
                {
                    //ports
                    Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_PhysicalAspect.Outputs["Route"] = port1;

                    Port port2 = new Port(OccurrenceConnection, part, "Lug", new Position(0, 0, pipeLugInputs.RodTakeOut), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_PhysicalAspect.Outputs["Lug"] = port2;


                    //Add Elbow geometry
                    if (HgrCompareDoubleService.cmpdbl(pipeLugInputs.Gap1, 0) == true) //For Single Lug  
                    {
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Rotate((Math.PI / 2), new Vector(0, 0, 1));
                        AddWeldingLug(pipeLugInputs, matrix, m_PhysicalAspect.Outputs, "pipeLugShape1");
                    }
                    else  //For Double Lug
                    {
                        matrix.Rotate((Math.PI / 2), new Vector(0, 0, 1));
                        matrix.Translate(new Vector( pipeLugInputs.Gap1 / 2 + pipeLugInputs.Thickness1 / 2,0, 0));
                        AddWeldingLug(pipeLugInputs, matrix, m_PhysicalAspect.Outputs, "pipeLugShape1");

                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Rotate((Math.PI / 2), new Vector(0, 0, 1));
                        matrix.Translate(new Vector( -pipeLugInputs.Gap1 / 2 - pipeLugInputs.Thickness1 / 2,0, 0));
                        AddWeldingLug(pipeLugInputs, matrix, m_PhysicalAspect.Outputs, "pipeLugShape2");

                        //Add Stiffener
                        if (HgrCompareDoubleService.cmpdbl(pipeLugInputs.StiffenerHeight, 0) == false && HgrCompareDoubleService.cmpdbl(pipeLugInputs.StiffenerLength, 0) == false)
                        {
                            symbolGeometryHelper.ActivePosition = new Position( -pipeLugInputs.Gap1 / 2,-pipeLugInputs.StiffenerLength / 2 , pipeLugInputs.RodTakeOut - pipeLugInputs.StiffenerHeight - pipeLugInputs.StiffenerOffset);
                            BusinessObject stiffener = symbolGeometryHelper.CreateBox(OccurrenceConnection, pipeLugInputs.Gap1, pipeLugInputs.StiffenerLength, pipeLugInputs.StiffenerHeight, 9);
                            m_PhysicalAspect.Outputs["pipeLugStiffener"] = stiffener;
                        }
                        //Add Pin
                        if (HgrCompareDoubleService.cmpdbl(pipeLugInputs.Pin1Diameter, 0) == false && HgrCompareDoubleService.cmpdbl(pipeLugInputs.Pin1Length, 0) == false)
                        {
                            symbolGeometryHelper = new SymbolGeometryHelper();
                            symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                            symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(1, 0, 0).GetOrthogonalVector());
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                           // matrix.Rotate((Math.PI / 2), new Vector(1, 0, 0));
                            Projection3d pipeLugPin = new Projection3d(symbolGeometryHelper.CreateCylinder(null, pipeLugInputs.Pin1Diameter / 2, pipeLugInputs.Pin1Length));
                            pipeLugPin.Transform(matrix);
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(-pipeLugInputs.Pin1Length / 2,0,   pipeLugInputs.RodTakeOut));
                            pipeLugPin.Transform(matrix);
                            m_PhysicalAspect.Outputs["pipeLugPin"] = pipeLugPin;
                        }
                    }
                }
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Welded Lug Attachment.");
                }
            }
        }
        #endregion

        #region "ICustomWeightCG Members"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                //System WCG Attributes
                Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];

                Double weight, cogX, cogY, cogZ;
                try
                {
                    weight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
                }
                catch
                {
                    weight = 0;
                }
                //Center of Gravity
                try
                {
                    cogX = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
                }
                catch
                {
                    cogX = 0;
                }
                try
                {
                    cogY = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
                }
                catch
                {
                    cogY = 0;
                }
                try
                {
                    cogZ = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
                }
                catch
                {
                    cogZ = 0;
                }

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of pipeLug"));
                    return;
                }
            }
        }
        #endregion


    }
}