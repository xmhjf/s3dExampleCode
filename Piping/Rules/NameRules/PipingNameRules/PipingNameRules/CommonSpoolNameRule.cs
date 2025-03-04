
//'*******************************************************************
//'  Copyright (C) 2004-2006 Intergraph Corporation.  All rights reserved.
//'
//'  Project: PipingNameRules
//'  Class:   CommonSpoolNameRule
//'
//'  Abstract: The file contains the Implementation for naming rule interface for Spools
//'
//'  Author:
//'
//'  History:
// '14/05/2015   chandrakala   CR-CP-269177  Route Pipe VB6 Rules need to be replaced  
//'******************************************************************
using Ingr.SP3D.Common.Middle;
using System;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Piping.Rules
{
    public class CommonSpoolNameRule : NameRuleBase
    {
        /// <summary>
        /// Computes the name for the given entity. 
        /// </summary>
        /// <param name="entity">The business object whose name is being computed.</param>
        /// <param name="parents">The parentsNaming parents of the business object, 
        /// i.e. other business objects which control the naming of the business object whose name is being computed.</param>
        /// <param name="activeEntity">The name rule active entity associated to the business object whose name is being computed.</param>
        public override void ComputeName(BusinessObject entity, ReadOnlyCollection<BusinessObject> parents, BusinessObject activeEntity)
        {
            if (entity == null)
            {
                throw new ArgumentNullException("entity");
            }
            if (parents == null)
            {
                throw new ArgumentNullException("The naming parents of the entity to be named are null");
            }
            if (activeEntity == null)
            {
                throw new ArgumentNullException("activeEntity");
            }
            try
            {
                string newName = string.Empty;
                BusinessObject pipeLine = null;
                if (parents.Count > 0) pipeLine = parents[0];
                // Get "TypeString" property value from IJDNamedItem interface
                string type = base.GetTypeString(entity);
                if (pipeLine != null)
                {
                    // Gets the Name  of the pipeLine                
                    string parentName = base.GetName(pipeLine);
                    // Get naming parents string from active entity
                    string namingParentsString = base.GetNamingParentsString(activeEntity);
                    //Name only needs to be recomputed and if the naming parent string is different from the existing one. 
                    if (!parentName.Equals(namingParentsString))
                    {
                        string location = string.Empty;
                        long counter = 0;
                        //Get the running count for the business object and location ID from the NameGeneratorService. 
                        base.GetCountAndLocationID(parentName, out counter, out location);
                        //Counter has to be padded with zeros for correct formatting. 
                        if (!string.IsNullOrEmpty(location))
                        {
                            newName = parentName + "_" + type + "-" + location + "-" + String.Format("{0:000}", counter);
                        }
                        else
                        {
                            newName = parentName + "_" + type + "-" + String.Format("{0:000}", counter);
                        }
                        base.SetNamingParentsString(activeEntity, newName);
                        // Set name on entity
                        base.SetName(entity, newName);
                    }
                }
            }
            catch (Exception)
            {
                throw new Exception("Unexpected error:  PipingNameRules.CommonSpoolNameRule.ComputeName");
            }

        }
        /// <summary>
        /// Gets the naming parents from naming rule.
        /// All the Naming Parents that need to participate in an objects naming are added here to the BusinessObject collection. 
        /// The parents added here are used in computing the name of the object in ComputeName() of the same interface. 
        /// Both these methods are called from naming rule semantic. 
        /// </summary>
        /// <param name="entity">BusinessObject for which naming parents are required.</param>
        /// <returns>Collection of BusinessObjects.</returns>
        public override Collection<BusinessObject> GetNamingParents(BusinessObject entity)
        {
            if (entity == null)
            {
                throw new ArgumentNullException("entity");
            }
            try
            {
                Collection<BusinessObject> parentsColl = new Collection<BusinessObject>();
                Spool spool = (Spool)entity;
                BusinessObject parentBo = null;
                try
                {
                    parentBo = (BusinessObject)spool.SpoolableObject;
                }
                catch (Exception)
                {
                    //do nothing                
                }
                parentsColl.Add(parentBo);
                return parentsColl;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
