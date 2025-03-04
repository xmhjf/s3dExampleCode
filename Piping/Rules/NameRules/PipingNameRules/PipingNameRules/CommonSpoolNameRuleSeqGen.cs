//'*******************************************************************
//'  Copyright (C) 2004-2006 Intergraph Corporation.  All rights reserved.
//'
//'  Project: PipingNameRules
//'  Class:   CommonSpoolNameRuleSeqGen
//'
//'  Abstract: The file contains the Implementation for naming rule interface for Spools
//'
// '14/05/2015   chandrakala   CR-CP-269177  Route Pipe VB6 Rules need to be replaced  
//'******************************************************************

using Ingr.SP3D.Common.Middle;
using System;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Piping.Rules
{
    public class CommonSpoolNameRuleSeqGen : NameRuleBase
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
                string parentName = string.Empty;
                // Get pipeline name
                if (parents.Count > 0) parentName = base.GetName(parents[0]);
                // Get "TypeString" property value from IJDNamedItem interface
                string type = base.GetTypeString(entity);
                PropertyValue sequenceID = entity.GetPropertyValue("IJSequence", "Id");
                PropertyValueString seqIDValue = (PropertyValueString)sequenceID;
                string seqId = seqIDValue.PropValue;
                //set name on entity
                base.SetName(entity, parentName + "_" + type + seqId);
            }
            catch (Exception)
            {
                throw new Exception("Unexpected error:  PipingNameRules.CommonSpoolNameRuleSeqGen.ComputeName");
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
