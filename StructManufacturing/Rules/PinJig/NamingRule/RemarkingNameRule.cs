//----------------------------------------------------------------------------------
//      Copyright (C) 2015 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   PinJig Remarking Name Rule. 
//
//      Author:  Suma Mallena
//
//      History:
//      April 21st, 2015   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Manufacturing.Middle.Services;
using System;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// Default PinJig Remarking Name Rule.
    /// </summary>
    public class RemarkingNameRule : ManufacturingSecondaryNameRuleBase
    {
        /// <summary>
        /// Provides the name for the Remarking Line.
        /// </summary>
        /// <param name="namingInfo">The Naming info.</param>
        public override void ComputeName(NamingInformation namingInfo)
        {
            try
            {
                RemarkingNamingInfomation remarkingNamingInfo = (RemarkingNamingInfomation)namingInfo;

                if (remarkingNamingInfo == null)
                    throw new ArgumentNullException("Input remarkingNamingInfo is empty.");

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //1. Get Inputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region GetInputs

                if ((remarkingNamingInfo.LongestRepresentedEntity == null) || (remarkingNamingInfo.SharedRepresentedEntities == null))
                    return;

                #endregion GetInputs

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing

                // Get Remarking Line Name
                string remarkingName = string.Empty;
                string partName = string.Empty;

                foreach (BusinessObject sharedEntity in remarkingNamingInfo.SharedRepresentedEntities)
                {
                    if ((sharedEntity is PlateSystem) || (sharedEntity is ProfileSystem))
                    {
                        partName = GetConnectedPartNames(sharedEntity, (PinJig)remarkingNamingInfo.ParentEntity);
                        //partName = GetPartsName(sharedEntity);//USE THIS LINE TO GET PART NAME BASED ON THE INTERSECTION.
                        if (remarkingName.Length > 0)
                        {
                            remarkingName = remarkingName + ", ";
                        }
                        remarkingName = remarkingName + partName;
                    }
                    else if (sharedEntity is Marking)
                    {
                        //TODO need to display name of Marking related part.
                        Marking marking = (Marking)sharedEntity;
                        if (remarkingName.Length > 0)
                        {
                            remarkingName = remarkingName + ", ";
                        }
                        remarkingName = remarkingName + marking.Name;
                    }
                    else
                    {
                        if (remarkingName.Length > 0)
                        {
                            remarkingName = remarkingName + ", ";
                        }
                        INamedItem namedItem = sharedEntity as INamedItem ;
                        if (namedItem != null)
                        {
                            remarkingName = remarkingName + namedItem.Name;
                        }
                        else
                        {
                            remarkingName = remarkingName + sharedEntity.ToString();
                        }
                    }
                }

                if (remarkingNamingInfo.ConsumedRepresentedEntities != null)
                {
                    int counter = 0;
                    foreach (BusinessObject consumedEntity in remarkingNamingInfo.ConsumedRepresentedEntities)
                    {
                        counter = counter + 1;
                        if (counter == 1)
                        {
                            remarkingName = remarkingName + ".  (Also ";
                        }

                        if ((consumedEntity is PlateSystem) || (consumedEntity is ProfileSystem))
                        {
                            partName = GetConnectedPartNames(consumedEntity, (PinJig)remarkingNamingInfo.ParentEntity);
                            //partName = GetPartsName(consumedEntity);//USE THIS LINE TO GET PART NAME BASED ON THE INTERSECTION.

                            remarkingName = remarkingName + ", " + partName;
                        }
                        else if (consumedEntity is Marking)
                        {
                            //TODO need to display name as related part name.
                            Marking markingLine = (Marking)consumedEntity;
                            remarkingName = remarkingName + ", " + markingLine.Name;
                        }
                        else
                        {
                            INamedItem namedItem = consumedEntity as INamedItem;
                            if (namedItem != null)
                            {
                                remarkingName = remarkingName + ", " + namedItem.Name;
                            }
                            else
                            {
                                remarkingName = remarkingName + ", " + consumedEntity.ToString();
                            }
                        }
                    }

                    if (counter > 0)
                    {
                        remarkingName = remarkingName + ").";
                    }
                }


                #endregion Processing

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs

                remarkingNamingInfo.Name = remarkingName;

                #endregion Set Outputs
            }
            catch (Exception e)
            {
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(true);
                string methodName = sf.GetMethod().Name;
                int lineNumber = sf.GetFileLineNumber();
                LogForToDoList(ToDoMessageTypes.ToDoMessageWarning, methodName, lineNumber, "SMCustomWarningMessages", 5022, "Call to PinJig Remarking Name Rule failed with the error." + e.Message);
            }
        }

        #region Private Methods
        /// <summary>
        /// Returns the connected part names.
        /// </summary>
        private string GetConnectedPartNames(BusinessObject representedEntity, PinJig pinjig)
        {
            ReadOnlyCollection<ISystemChild> systemChildren = null;
            Collection<BusinessObject> childParts = new Collection<BusinessObject>();
            ISurface remarkingSurface = pinjig.RemarkingSurface;
            Collection<BusinessObject> connectedObjects = new Collection<BusinessObject>();
            string remarkingName = string.Empty;

            if (representedEntity is PlateSystem)
            {
                Plate plateSystem = (Plate)representedEntity;
                systemChildren = plateSystem.SystemChildren;
                for (int index = 0; index < systemChildren.Count; index++)
                {
                    if (systemChildren[index] is PlatePart)
                    {
                        childParts.Add((PlatePartBase)systemChildren[index]);
                    }
                }
            }
            else if (representedEntity is ProfileSystem)
            {
                Profile profileSystem = (Profile)representedEntity;
                systemChildren = profileSystem.SystemChildren;
                for (int index = 0; index < systemChildren.Count; index++)
                {
                    if (systemChildren[index] is ProfilePart)
                    {
                        childParts.Add((ProfilePart)systemChildren[index]);
                    }
                }
            }

            double distance = 0.0;
            Position position = null;
            for (int index = 0; index < childParts.Count; index++)
            {
                GeometryService.GetMinimumDistance((BusinessObject)remarkingSurface, childParts[index], out position, out distance);
                if (distance < 0.01)
                {
                    connectedObjects.Add(childParts[index]);
                }
            }
            for (int index = 0; index < connectedObjects.Count; index++)
            {
                if (remarkingName.Length > 0)
                {
                    remarkingName = remarkingName + ", ";
                }
                INamedItem namedItem = (INamedItem)connectedObjects[index];
                remarkingName = remarkingName + namedItem.Name;
            }
            return remarkingName;
        }

        /// <summary>
        /// Returns the part names without considering the intersection with Pinjig surface.
        /// Logic:
        ///   Get the parts from the system
        ///   Get the names of the parts
        ///   Get the last separator '-' position
        ///   Compare all the part names till the last '-'
        ///   If names are different, concatenate all parts names
        /// </summary>
        private string GetPartsName(BusinessObject representedEntity)
        {
            ReadOnlyCollection<ISystemChild> systemChildren = null;
            Collection<BusinessObject> childParts = new Collection<BusinessObject>();
            string remarkingName = string.Empty;
            if (representedEntity is PlateSystem)
            {
                Plate plateSystem = (Plate)representedEntity;
                systemChildren = plateSystem.SystemChildren;
                for (int index = 0; index < systemChildren.Count; index++)
                {
                    if (systemChildren[index] is PlatePart)
                    {
                        childParts.Add((PlatePartBase)systemChildren[index]);
                    }
                }
            }
            else
            {
                if (representedEntity is ProfileSystem)
                {
                    Profile profileSystem = (Profile)representedEntity;
                    systemChildren = profileSystem.SystemChildren;
                    for (int index = 0; index < systemChildren.Count; index++)
                    {
                        if (systemChildren[index] is ProfilePart)
                        {
                            childParts.Add((ProfilePart)systemChildren[index]);
                        }
                    }
                }
            }
            INamedItem namedItem = (INamedItem)childParts[0];
            remarkingName = namedItem.Name;
            int position = 0;
            string commonString = string.Empty;
            if (childParts.Count > 0)
            {
                position = remarkingName.LastIndexOf("-");
                if (position != -1)
                {
                    commonString = remarkingName.Substring(0, position);

                    for (int index = 1; index < childParts.Count; index++)
                    {
                        if (remarkingName.IndexOf(commonString) == 0)
                        {
                            commonString = string.Empty;
                            break;
                        }
                    }
                }
                if (commonString == string.Empty)
                {
                    for (int index = 1; index < childParts.Count; index++)
                    {
                        if (commonString.Length > 0)
                        {
                            commonString = commonString + ", ";
                        }
                        namedItem = (INamedItem)childParts[index];
                        commonString = commonString + namedItem.Name;
                    }
                }
            }
            return remarkingName;
        }

        #endregion Private Methods
    }
}
