//----------------------------------------------------------------------------------
//      Copyright (C) 2015 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   BlockMarkDefinition is a .NET Symbol Definition for Block Marks.
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
using Ingr.SP3D.Manufacturing.Middle;
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
    public class BlockMarkDefinition : GenericMarkRuleBase
    {
        #region Definition of Inputs

        /// <summary>
        /// PinJig Input
        /// </summary>
        [InputObject(1, SymbolsConstants.PinJigInput, SymbolsConstants.PinJigInput)]
        public InputObject pinJig;

        /// <summary>
        /// Supported Plates By User Input Input
        /// </summary>
        [InputObject(2, SymbolsConstants.SupportedPlatesByUserInput, SymbolsConstants.SupportedPlatesByUserInput)]
        public InputObject supportedPlatesByUser; 

        /// <summary>
        /// Offset Value Input
        /// </summary>
        [InputDouble(3, SymbolsConstants.OffsetValueInput, SymbolsConstants.OffsetValueInput, 0.0)]
        public InputDouble offsetValue;

        /// <summary>
        /// Supported Plates By Type Input
        /// </summary>
        [InputDouble(4, SymbolsConstants.SupportedPlatesByTypeInput, SymbolsConstants.SupportedPlatesByTypeInput, 0.0)] 
        public InputDouble supportedPlatesByType;

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
        /// Computes the block mark for the given inputs.
        /// </summary>
        protected override void ConstructOutputs()
        {
            try
            {
          
                if ((this.pinJig.Value == null))
                {
                    //The symbol input is not set yet.
                    return;
                }

                if (StructHelper.AreEqual(this.offsetValue.Value, 0.0) ||
                    StructHelper.AreEqual(this.supportedPlatesByType.Value, 0.0))                
                {
                    //The symbol inputs OffsetValue and SuppPlatesByType are not set yet
                    return;
                }

                Dictionary<string, ComplexString3d> outputs = null;
                outputs = CreateOutputs((PinJig)this.pinJig.Value, this.offsetValue.Value, (int)this.supportedPlatesByType.Value, (BusinessObject) this.supportedPlatesByUser.Value);

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
                base.WriteToErrorLog(methodName, lineNumber, "Call to BlockMarkDefinition rule failed with error : ", e.Message);

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
                if (inputs[3] != null)
                {
                    string platesType = Convert.ToString(inputs[3]) ;
                    
                    if( platesType == "User Selection")
                    {
                        entitiesToMark.Add((BusinessObject)inputs[4]);
                    }
                    else
                    {
                        int plateType = 0;
                        if( platesType == "Transversal")
                        {
                            plateType = 1;
                        }
                        else if( platesType == "Longitudinal")
                        {
                            plateType = 1;
                        }
                        else// if( platesType == "Transversal and Longitudinal")
                        {
                            plateType = 1;
                        }

                        PinJig pinJig = (PinJig) inputs[1];
                        foreach( PlatePartBase basePlatePart in pinJig.SupportedPlates )
                        {
                            Collection<PlatePartBase> connectedPlates = GetPlateConnectedObjectsByType(basePlatePart, plateType);
                            if (connectedPlates != null)
                            {
                                foreach (PlatePartBase connectedPlate in connectedPlates)
                                {
                                    entitiesToMark.Add(connectedPlate);
                                }
                            }
                        }
                    }                    
                }

            }
            catch (Exception e)
            {
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(true);
                string methodName = sf.GetMethod().Name;
                int lineNumber = sf.GetFileLineNumber();
                base.WriteToErrorLog(methodName, lineNumber, "Call to BlockMarkDefinition rule failed with error : ", e.Message);
            }

            return new ReadOnlyCollection<BusinessObject>(entitiesToMark);
        }

        /// <summary>
        /// Returns the collection of information for all the inputs of the generic mark.
        /// </summary>
        /// <param name="genericMarkingInformation">The generic marking information for getting the configuration data.</param> 
        public override ReadOnlyCollection<GenericMarkInputInformation> GetInputsInformation(GenericMarkingInformation genericMarkingInformation)
        {
            Collection<GenericMarkInputInformation> inputInformationCollection = new Collection<GenericMarkInputInformation>();

            const string IID_IJPinJig = "{FE221533-5879-11D5-B86E-0000E2300200}";
            const string IID_IJPlatePart = "{780F26C2-82E9-11D2-B339-080036024603}";

            try
            {
                //if (genericMarkingInformation == null)
                //    throw new CmnArgumentNullException("genericMarkingInformation.");

                //PinJig
                GenericMarkInputInformation inputInformation = new GenericMarkInputInformation();
                inputInformation.InputType = InputType.Object;
                inputInformation.Prompt = "Select a Pin Jig";
                inputInformation.DisplayName = "Pin Jig";
                inputInformation.DefaultValue = IID_IJPinJig;
                inputInformation.IsOptional = false;

                inputInformationCollection.Add(inputInformation);

                //Offset Value
                inputInformation = new GenericMarkInputInformation();
                inputInformation.InputType = InputType.DimensionedValue;
                inputInformation.Prompt = "Enter Offset Value";
                inputInformation.DisplayName = "Offset Value";
                inputInformation.DefaultValue = 2.5;
                inputInformation.IsOptional = false;

                inputInformationCollection.Add(inputInformation);

                //Supported Plates by Type
                inputInformation = new GenericMarkInputInformation();
                inputInformation.InputType = InputType.CodeListedValue;
                inputInformation.Prompt = "Select Supported Plates by Type";
                inputInformation.DisplayName = "Supported Plates by Type";
                inputInformation.DefaultValue = "StrMfgMarkingConnPlatesType:1";
                inputInformation.IsOptional = false;

                inputInformationCollection.Add(inputInformation);

                //Supported Plates by User
                inputInformation = new GenericMarkInputInformation();
                inputInformation.InputType = InputType.MultipleObject;
                inputInformation.Prompt = "Select Plates manually: Allowed only if Supported Plates by Type is 'User Selection'";
                inputInformation.DisplayName = "Supported Plates by User";
                inputInformation.DefaultValue = IID_IJPlatePart + " AND " + " [MfgGenericMarkSymbol.CustomFilterHelper,IsConnectedToPinJigSuppPlates] ";
                inputInformation.IsOptional = false;
                inputInformationCollection.Add(inputInformation);

            }
            catch (Exception e)
            {
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(true);
                string methodName = sf.GetMethod().Name;
                int lineNumber = sf.GetFileLineNumber();
                base.WriteToErrorLog(methodName, lineNumber, "Call to BlockMarkDefinition rule failed with error : ", e.Message);
            }

            return new ReadOnlyCollection<GenericMarkInputInformation>(inputInformationCollection);
        }

        #endregion override properties and methods

        #region Private Methods

        /// <summary>
        /// Creates the symbol outputs.
        /// </summary>
        /// <param name="plateSystem">The plate system.</param>
        /// <param name="offsetValue">The offset value.</param>
        /// <param name="supportedPlatesByType">The supported plates selected by the type.</param>
        /// <param name="supportedPlatesByUser">The supported plates selected by the user.</param>
        private Dictionary<string, ComplexString3d> CreateOutputs(PinJig plateSystem, double offsetValue, int supportedPlatesByType, BusinessObject supportedPlatesByUser)
        {
            Dictionary<string, ComplexString3d> outputs = new Dictionary<string, ComplexString3d>();
            return outputs;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="basePlatePart"></param>
        /// <param name="platesType"></param>
        /// <returns></returns>
        private Collection<PlatePartBase> GetPlateConnectedObjectsByType(PlatePartBase basePlatePart, int platesType) 
        {
            return null;
        }

        #endregion Private Methods

    }
}