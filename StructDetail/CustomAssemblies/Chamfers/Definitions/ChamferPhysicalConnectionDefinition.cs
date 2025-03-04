//------------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  ChamferPhysicalConnectionDefinition.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMChamferRules.dll
//  Original Class Name: ‘ChamferPhysConnDef’ in VB content
//
//Abstract
//     ChamferPhysicalConnectionDefinition is a .NET definition rule which is defining the custom definition rule. 
//  This class derives from ChamferCustomAssemblyDefinition. 
//--------------------------------------------------------------------------------------------------------------
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.Structure.Middle;
using System;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Implementation of Chamfer .NET custom assembly definition class. ChamferPhsicalConnection defintion is responsible
    /// to create its only assembly output i.e., physical connection.
    /// </summary>
    [SymbolVersion("1.0.0.0")]
    [DefaultLocalizer(ChamfersResourceIds.ChamfersResources, ChamfersResourceIds.ChamfersAssembly)]
    public class ChamferPhysicalConnectionDefinition : ChamferCustomAssemblyDefinition
    {
        //===================================================================================================================
        //DefinitionName/ProgID of this symbol is "Chamfers,Ingr.SP3D.Content.Structure.ChamferPhysicalConnectionDefinition"
        //===================================================================================================================
        #region Definitions of assembly outputs

        //Creates a Chamfer Physical Connection
        [AssemblyOutput(1, DetailingCustomAssembliesConstants.ChamferPhysConn1)]
        public AssemblyOutput physicalConnectionAssemblyOutput;

        #endregion Definitions of assembly outputs

        #region Public override functions and methods

        /// <summary>
        /// Construct and re-evaluate the custom assembly output. It creates the physicalconnection if not exists 
        /// and copies the default answers on the physical connection from its parent.
        /// </summary>
        public override void EvaluateAssembly()
        {
            try
            {
                //Validating inputs
                base.ValidateChamferInputs();

                //Check if any ToDo message of type error has been created during the validation of inputs
                if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                {
                    //ToDo list is created while validating the inputs
                    return;
                }

                //the following call is needed if the parent definition item is in vb6,
                //in case of .Net CAD, the constructor call itself taking care of adding geometry to the part
                base.AddChamferGeometry();

                Ingr.SP3D.Interop.StructDetailUtil.StructDetailHelper oHlpr = new Ingr.SP3D.Interop.StructDetailUtil.StructDetailHelper();
                bool isPartialDetailed = false;

                if (!(base.ChamferedPort.Connectable is MemberPart ))
                {
                    isPartialDetailed = (oHlpr.IsPartialDetailed(Ingr.SP3D.Common.Middle.Services.Hidden.COMConverters.ConvertBOToCOMBO((Ingr.SP3D.Common.Middle.BusinessObject)base.ChamferedPort.Connectable)));                    
                }

                if (!(base.DrivesChamferedPort.Connectable is MemberPart))
                {
                    isPartialDetailed = oHlpr.IsPartialDetailed(Ingr.SP3D.Common.Middle.Services.Hidden.COMConverters.ConvertBOToCOMBO((Ingr.SP3D.Common.Middle.BusinessObject)base.DrivesChamferedPort.Connectable)) || isPartialDetailed;
                }
                
                oHlpr = null;

                if (isPartialDetailed)
                {
                    return;
                }

                PhysicalConnection physicalConnection;
                //check if output is already created, if not create the physical connection
                if (this.physicalConnectionAssemblyOutput.Output == null)
                {
                    //there is no condition to control the physical connection creation, so always create the PC if not available
                    physicalConnection = CreatePhysicalConnection(DetailingCustomAssembliesConstants.ButtWeld);
                    if (physicalConnection != null)
                    {
                        this.physicalConnectionAssemblyOutput.Output = physicalConnection;
                    }
                    else
                    {
                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(ChamfersResourceIds.CreatePhysicalConnectionError,
                        "Error while creating the physical connection."));
                    }
                }
                else
                {
                    //get the existing output
                    physicalConnection = (PhysicalConnection)this.physicalConnectionAssemblyOutput.Output;
                }

                if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                {
                    //ToDo list is created with error type hence stop computation
                    return;
                }

                //now, copy default answers to its child
                this.SetAnswers(physicalConnection);
            }
            catch(Exception)
            {
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(ChamfersResourceIds.ToDoEvaluateAssembly,
                        "Unexpected error while evaluating the chamfer physical connection definition."));
                }
            }
        }

        #endregion Public override functions and methods

        #region Private functions and methods

        /// <summary>
        /// Sets the answer of the Chamfer selector questions to its assembly output.
        /// </summary>
        /// <param name="physicalConnection">The physical connection.</param>
        private void SetAnswers(PhysicalConnection physicalConnection)
        {
            try
            {
                Feature chamfer = (Feature)base.Occurrence;

                //Set the answer of the parent's (Chamfer) ChamferType question on the PhysicalConnection.
                base.SetAnswer(physicalConnection, DetailingCustomAssembliesConstants.ChamferType);

                //Get the answer of Shipyard_FromAssyConn question from the Chamfer and then set the same for the Shipyard question on the PhysicalConnection 
                PropertyValue shipyardName = chamfer.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.Shipyard);
                if (shipyardName != null)
                {
                    base.SetAnswer(physicalConnection, DetailingCustomAssembliesConstants.Shipyard_FromAssyConn, shipyardName);
                }

                //Get the answer of WeldingType_FromAssyConn question from the Chamfer and then set the same for the WeldingType question on the PhysicalConnection
                PropertyValue weldingType = chamfer.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.WeldingType);
                if (weldingType != null)
                {
                    base.SetAnswer(physicalConnection, DetailingCustomAssembliesConstants.WeldingType_FromAssyConn, weldingType);
                }
            }
            catch
            {
                //create generic ToDo message only when there is no specific ToDo messages are created.
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(ChamfersResourceIds.ErrCopyAnswers,
                        "Unexpected error while copying answers of chamfer to the physical connection."));
                }
            }
        }

        #endregion Private functions and methods
    }
}
