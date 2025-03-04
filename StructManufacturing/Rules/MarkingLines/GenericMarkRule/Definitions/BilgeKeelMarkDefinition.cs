//----------------------------------------------------------------------------------
//      Copyright (C) 2015 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   BilgeKeelMarkDefinition is a .NET Symbol Definition for Bilge Keel Marks.
//                 
//
//      Author:  
//
//      History:
//      November 5th, 2015   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using Ingr.SP3D.Common.Exceptions;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Manufacturing
{

    /// <summary>
    /// BilgeKeel Mark Symbol Definition.
    /// </summary>
    [SymbolVersion("1.0.0.0")]
    public class BilgeKeelMarkDefinition : GenericMarkRuleBase
    {
        #region Definition of Inputs

        /// <summary>
        /// Plate System Input
        /// </summary>
        [InputObject(1, SymbolsConstants.PlateSystemInput, SymbolsConstants.PlateSystemInput)]
        public InputObject plateSystem;

        /// <summary>
        /// Molded Form System Input
        /// </summary>
        [InputObject(2, SymbolsConstants.MoldedFormSystemInput, SymbolsConstants.MoldedFormSystemInput)]
        public InputObject connectedSystem;

        /// <summary>
        /// Reference Surface Input
        /// </summary>
        [InputDouble(3, SymbolsConstants.ReferenceSurfaceInput, SymbolsConstants.ReferenceSurfaceInput, 0.0)]
        public InputDouble referenceSurface;

        /// <summary>
        /// Marking Side Input
        /// </summary>
        [InputDouble(4, SymbolsConstants.MarkingSideInput, SymbolsConstants.MarkingSideInput, 0.0)]
        public InputDouble markingSide;        

        #endregion

        #region Define aspect
        /// <summary>
        /// The simple physical
        /// </summary>
        [Aspect(SymbolsConstants.GenericMarkAspect, SymbolsConstants.GenericMarkAspect, AspectID.SimplePhysical)]
        public AspectDefinition simplePhysical;        

        #endregion
    
        #region Override properties and methods

        /// <summary>
        /// Computes the Bilge Keel mark for the given inputs.
        /// </summary>
        protected override void ConstructOutputs()
        {
            try
            {
                if ((this.plateSystem.Value == null))
                {
                    //The symbol input is not set yet.
                    return;
                }
               
                if( StructHelper.AreEqual(this.referenceSurface.Value, 0.0) ||
                    StructHelper.AreEqual(this.markingSide.Value, 0.0))
                {
                    //The symbol inputs MarkingSide and Reference Surface are not set yet.
                    return;
                }

                Dictionary<string, ComplexString3d> outputs = null;
                outputs = CreateOutputs((Plate)this.plateSystem.Value, this.connectedSystem.Value, (int)this.markingSide.Value, (int)this.referenceSurface.Value);

                if (outputs != null)
                {
                    foreach (KeyValuePair<string, ComplexString3d> output in outputs)
                    {
                        simplePhysical.Outputs.Add(output.Key, output);
                    }
                }
            }
            catch (Exception e) // General Unhandled exception 
            {
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(true);
                string methodName = sf.GetMethod().Name;
                int lineNumber = sf.GetFileLineNumber();
                base.WriteToErrorLog(methodName, lineNumber, "Call to BilgeKeelMarkDefinition rule failed with error : ", e.Message);

            }
        }

        /// <summary>
        /// Returns the parts or systems to be marked from the inputs.
        /// </summary>
        /// <param name="genericMarkingInformation">The generic marking information for getting the configuration data.</param>
        /// <param name="inputs">Collection of inputs.</param>
        public override ReadOnlyCollection<BusinessObject> GetEntitiesToMark(GenericMarkingInformation genericMarkingInformation, ReadOnlyCollection<object> inputs)
        {
            Collection<BusinessObject> entitiesToMark = new Collection<BusinessObject>();
         
            try
            {
                //if (genericMarkingInformation == null)
                //    throw new CmnArgumentNullException("genericMarkingInformation.");

                if (inputs == null)
                    throw new CmnArgumentNullException("inputs");

                // The first input in the symbol is plate system
                if (inputs[0] != null)
                {
                    Plate plateSystem = inputs[0] as Plate;
                    if (plateSystem != null)
                    {
                        entitiesToMark.Add(plateSystem);
                    }
                }

            }
            catch (Exception e)
            {         
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(true);
                string methodName = sf.GetMethod().Name;
                int lineNumber = sf.GetFileLineNumber();
                base.WriteToErrorLog(methodName, lineNumber, "Call to BilgeKeelMarkDefinition rule failed with error : ", e.Message);
            }

            return new ReadOnlyCollection<BusinessObject> (entitiesToMark);
        }

        /// <summary>
        /// Returns the collection of information for all the inputs of the generic mark.
        /// </summary>
        /// <param name="genericMarkingInformation">The generic marking information for getting the configuration data.</param> 
        public override ReadOnlyCollection<GenericMarkInputInformation> GetInputsInformation(GenericMarkingInformation genericMarkingInformation)
        {
            Collection<GenericMarkInputInformation> inputInformationCollection = new Collection<GenericMarkInputInformation>();

            const string IID_IJPlateSystem = "{E0B23CD4-7CEB-11d3-B351-0050040EFC17}";
            const string IID_IJStiffenerSystem = "{E0B23CD5-7CEB-11d3-B351-0050040EFC17}";

            try
            {
                //if (genericMarkingInformation == null)
                //    throw new CmnArgumentNullException("genericMarkingInformation.");

                //PlateSystem
                GenericMarkInputInformation inputInformation = new GenericMarkInputInformation();
                inputInformation.InputType = InputType.Object ;
                inputInformation.Prompt = "Select a Plate System";
                inputInformation.DisplayName = "Plate System";
                inputInformation.DefaultValue = IID_IJPlateSystem + " AND " + " NOT [STFilterFunctions.StructFilterFunctions,ResultOfSplit] ";
                inputInformation.IsOptional = false;

                inputInformationCollection.Add(inputInformation);

                //PlateSystem/ProfileSystem
                inputInformation = new GenericMarkInputInformation();
                inputInformation.InputType = InputType.Object;
                inputInformation.Prompt = "Select a connected Plate or Profile System";
                inputInformation.DisplayName = "Plate/Profile System";
                inputInformation.DefaultValue = IID_IJPlateSystem + " AND " + " [MfgGenericMarkSymbol.CustomFilterHelper,IsConnectedToPlateSystem] " + " AND " + " NOT [STFilterFunctions.StructFilterFunctions,ResultOfSplit] " +
                                         " OR " + IID_IJStiffenerSystem + " AND " + " [MfgGenericMarkSymbol.CustomFilterHelper,IsConnectedToPlateSystem] " + " AND " + " NOT [STFilterFunctions.StructFilterFunctions,ResultOfSplit] ";
                inputInformation.IsOptional = false;

                inputInformationCollection.Add(inputInformation);

                //Reference Surface
                inputInformation = new GenericMarkInputInformation();
                inputInformation.InputType = InputType.CodeListedValue;
                inputInformation.Prompt = "Select connected part's reference surface";
                inputInformation.DisplayName = "Reference Surface";
                inputInformation.DefaultValue = "MfgMarkingLinesSide:2110"; //Web Left
                inputInformation.IsOptional = false;

                inputInformationCollection.Add(inputInformation);

                //Marking Side
                inputInformation = new GenericMarkInputInformation();
                inputInformation.InputType = InputType.CodeListedValue;
                inputInformation.Prompt = "Marking Side";
                inputInformation.DisplayName = "Marking Side";
                inputInformation.DefaultValue = "MfgMarkingLinesSide:1111"; //Molded Side
                inputInformation.IsOptional = false;
                inputInformationCollection.Add(inputInformation);

            }
            catch (Exception e)
            {
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(true);
                string methodName = sf.GetMethod().Name;
                int lineNumber = sf.GetFileLineNumber();
                base.WriteToErrorLog(methodName, lineNumber, "Call to BilgeKeelMarkDefinition rule failed with error : ", e.Message);
            }

            return new ReadOnlyCollection<GenericMarkInputInformation>(inputInformationCollection);
        }
        
        #endregion override properties and methods

        #region Private Methods

        /// <summary>
        /// Creates the symbol outputs.
        /// </summary>
        /// <param name="plateSystem">The plate system.</param>
        /// <param name="connectedSystem">The connected plate system.</param>
        /// <param name="markingSide">The marking side.</param>
        /// <param name="referenceSurface">The reference surface.</param>
        private Dictionary<string, ComplexString3d> CreateOutputs(Plate plateSystem, BusinessObject connectedSystem, int markingSide, int referenceSurface)
        {
            Dictionary<string, ComplexString3d> outputs = new Dictionary<string, ComplexString3d>();
            return outputs;
        }

        #endregion Private Methods

    }
}