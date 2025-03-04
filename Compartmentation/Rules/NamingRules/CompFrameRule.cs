/*'***************************************************************************
'  Copyright (C) 2015, Intergraph Corporation.  All rights reserved.
'
'  Project: NamingRules
'
'  Abstract: The file contains an implementation of the Frame naming rule
'            for the Compartment objectin Compartmentation UE.
'           It finds out the bounding Frames of the range of the compartment and
'           assigns the names of the compartment as per these frames.
'
'  History:
'  Sravanya            10th March 2015               Creation
'****************************************************************************/


using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Compartmentation.Middle;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace Ingr.SP3D.Content.Compartmentation
{
    /// <summary>
    ///Enumerated values to be used in combinations for getting the Reference Position
    /// </summary>
    public enum ReferencePosition
    {
        /// <summary>
        /// Above
        /// </summary>
        Above = 6,

        /// <summary>
        ///Aft
        /// </summary>
        Aft = 2,

        /// <summary>
        /// Below
        /// </summary>
        Below = 7,

        /// <summary>
        /// Fore
        /// </summary>
        Fore = 3,

        /// <summary>
        ///OnReference
        /// </summary>
        OnReference = 1,

        /// <summary>
        /// PortSide
        /// </summary>
        PortSide = 4,

        /// <summary>
        ///StarBoard
        /// </summary>
        StarBoard = 5,

        /// <summary>
        /// Undefined
        /// </summary>
        Undefined = 0
    }

    public class CompFrameRule : NameRuleBase
    {
        private const string strFormat = "{0:0000}";

        /// <summary>
        /// Creates a name for the object passed in. The name is based on the parents name and object name. It is assumed that all Naming Parents and the Object implement INamedItem.
        /// The Naming Parents are added in AddNamingParents() of the same interface Both these methods are called from the naming rule semantic. Notes: ZoneName  = Zone description + Unique Index
        /// </summary>
        /// <param name="entity">The business object whose name is being computed. </param>
        /// <param name="parents">Naming parents of the business object, 
        /// i.e. other business objects which control the naming of the business object whose name is being computed.</param>
        /// <param name="activeEntity">The name rule active entity associated to the business object whose name is being computed.</param>
        public override void ComputeName(BusinessObject entity, ReadOnlyCollection<BusinessObject> parents, BusinessObject activeEntity)
        {
            try
            {
                if (entity == null)
                {
                    throw new ArgumentNullException("entity");
                }
                if (activeEntity == null)
                {
                    throw new ArgumentNullException("activeEntity");
                }  

                string strName = string.Empty , strLocation = string.Empty;
                long count;
                string boundingFrameName= string.Empty; //'define fixed-width number field
                BusinessObject boundingFrame = null;

                ReferencePosition eReferencePosition = getReferencePosition(entity);

                if (eReferencePosition == ReferencePosition.PortSide)
                {
                    strName = "Port";
                }
                else if (eReferencePosition == ReferencePosition.StarBoard)
                {
                    strName = "StarBoard";
                }
                else
                {
                    strName = "Center";
                }

                RuleHelper ruleHelper = new RuleHelper();
                Frames frames = ruleHelper.GetFrames(entity);

                // 'For X Axis

                boundingFrame = frames.xLow;

                if (boundingFrame!=null)
                {
                    boundingFrameName = boundingFrameName + boundingFrame.GetPropertyValue("IJNamedItem","Name");
                }
                    
                boundingFrame = frames.xHigh;

                if (boundingFrame != null)
                {
                    if (frames.xLow!= null)
                    {
                        boundingFrameName = boundingFrameName + "-" +boundingFrame.GetPropertyValue("IJNamedItem", "Name");
                    }
                    else
                    {
                        boundingFrameName = boundingFrameName + boundingFrame.GetPropertyValue("IJNamedItem", "Name");
                    }
                }

                // 'For Y Axis

                boundingFrame = frames.yLow;

                if (boundingFrame != null)
                {
                    if (frames.xLow != null || frames.xHigh != null)
                    {
                        boundingFrameName = boundingFrameName + "::" + boundingFrame.GetPropertyValue("IJNamedItem", "Name");
                    }
                    else
                    {
                        boundingFrameName = boundingFrameName + boundingFrame.GetPropertyValue("IJNamedItem", "Name");
                    }
                }

                boundingFrame = frames.yHigh;

                if (boundingFrame != null)
                {
                    if (frames.yLow != null)
                    {
                        boundingFrameName = boundingFrameName + "-" + boundingFrame.GetPropertyValue("IJNamedItem", "Name");
                    }
                    else
                    {
                        boundingFrameName = boundingFrameName + boundingFrame.GetPropertyValue("IJNamedItem", "Name");
                    }
                }

                // 'For Z Axis

                boundingFrame = frames.zLow;

                if (boundingFrame != null)
                {
                    if (frames.yLow != null || frames.yHigh!=null)
                    {
                        boundingFrameName = boundingFrameName + "::" + boundingFrame.GetPropertyValue("IJNamedItem", "Name");
                    }
                    else
                    {
                        boundingFrameName = boundingFrameName + boundingFrame.GetPropertyValue("IJNamedItem", "Name");
                    }
                }

                boundingFrame = frames.zHigh;

                if (boundingFrame != null)
                {
                    if (frames.zLow != null)
                    {
                        boundingFrameName = boundingFrameName + "-" + boundingFrame.GetPropertyValue("IJNamedItem", "Name");
                    }
                    else
                    {
                        boundingFrameName = boundingFrameName + boundingFrame.GetPropertyValue("IJNamedItem", "Name");
                    }
                }

                if (boundingFrameName!=string.Empty)
                {
                    strName = strName +"_" + boundingFrameName;
                }
                else
                {
                    strName = strName + boundingFrameName;
                }

                //Returns the number of occurrence of a string in addtion to the LocationID
                GetCountAndLocationID(strName, out count, out strLocation);

                //Add LocationID, if available
                if (strLocation!=string.Empty)
                {
                    strName = strName +"-" + strLocation + "-" + string.Format(strFormat,count);
                }
                else
                {
                    strName = strName + "-" + string.Format(strFormat, count);
                }

                entity.SetPropertyValue(strName ,"IJNamedItem", "Name");
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("FrameRule.ComputeName: Error encountered (" + e.Message + ")");
            }
        }

        /// <summary>
        ///   All the Naming Parents that need to participate in an objects naming are added here to the collection. Dummy function which does nothing
        /// </summary>
        /// <param name="entity">The business object whose name is being computed. </param>
        public override Collection<BusinessObject> GetNamingParents(BusinessObject entity)
        {
            Collection<BusinessObject> oRetColl = new Collection<BusinessObject>();
            try
            {
                if (entity == null)
                {
                    throw new ArgumentNullException("entity");
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("FrameRule.GetNamingParents: Error encountered (" + e.Message + ")");
            }

            return oRetColl;
        }

        /// <summary>
        /// gets the reference position for the given range.
        /// </summary>
        /// <param name="range">range of the entity is passed </param>
        /// <returns>string</returns>
        private ReferencePosition getReferencePosition(BusinessObject entity)
        {
            ReferencePosition ePosition = ReferencePosition.Undefined;
            try
            {
                Compartment compartment = (Compartment)entity;
                RangeBox rangeBox = compartment.Range;

                if (rangeBox.Low.Y >= 0 && rangeBox.High.Y >= 0)
                {
                    ePosition = ReferencePosition.PortSide;
                }
                else if (rangeBox.Low.Y <= 0 && rangeBox.High.Y <= 0)
                {
                    ePosition = ReferencePosition.StarBoard;
                }
                return ePosition;
            }

            catch (Exception e)
            {
                throw e;
            }
        }

    }
}
