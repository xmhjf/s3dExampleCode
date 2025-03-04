using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Planning
{
    /// <summary>
    /// SeqBasePlateProfile rule class Implements PlnAssemblyRangeRuleBase
    /// </summary>
    public class SeqBasePlateProfile : PlnAssemblySequenceRuleBase
    {
        /// <summary>
        ///  Sequence base plates and profiles
        /// </summary>
        /// <param name="sequenceData"></param>
        public override void Evaluate(SequenceData sequenceData)
        {
            #region Input Error Handling

            if (sequenceData == null)
            {
                throw new ArgumentNullException("sequenceData Information is null");
            }

            #endregion //Input Error Handling

            try
            {
                int basePlateCount = 0;
                Matrix4X4 oViewMatrix = null;
                object[] oUnrelatedPart = null;

                BusinessObject assemblyBase = sequenceData.Assembly;
                Collection<BusinessObject> assyChildren = new Collection<BusinessObject>();
                Collection<PLATE_INFO> basePlateInfoColl = new Collection<PLATE_INFO>();

                foreach (BusinessObject children in sequenceData.AssemblyChildren)
                {
                    assyChildren.Add(children);
                }

                // Specify parts that will be sequenced
                int sequenceFilter = SEQUENCE_FILTER_BASEPLATE_NON_HULL | SEQUENCE_FILTER_STIFFENING_PROFILE;
                int nUnrelatedPartCount = 0;

                CollectBasePlateInfo(assemblyBase, assyChildren, sequenceFilter, basePlateInfoColl, ref basePlateCount, ref oViewMatrix, ref nUnrelatedPartCount, ref oUnrelatedPart);

                // Sequence all plate parts and profile parts for each deck                

                for (int basePlateIndex = 0; basePlateIndex <= basePlateCount - 1; basePlateIndex++)
                {
                    SequenceParts(assemblyBase, basePlateInfoColl[basePlateIndex], oViewMatrix, SequenceType.SeqBasePlateProf);
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
        }      
 
    }
}
