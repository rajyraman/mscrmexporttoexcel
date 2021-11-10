// PROJECT : MsCrmTools.ViewLayoutReplicator
// This project was developed by Tanguy Touzard
// CODEPLEX: http://xrmtoolbox.codeplex.com
// BLOG: http://mscrmtools.blogspot.com

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace MsCrmTools.ViewLayoutReplicator.Helpers
{
    /// <summary>
    /// Helps to interact with Crm views
    /// </summary>
    class ViewHelper
    {
        #region Constants

        public const int VIEW_BASIC = 0;
        public const int VIEW_ADVANCEDFIND = 1;
        public const int VIEW_ASSOCIATED = 2;
        public const int VIEW_QUICKFIND = 4;
        public const int VIEW_SEARCH = 64;

        #endregion

        /// <summary>
        /// Retrieve the list of views for a specific entity
        /// </summary>
        /// <param name="selectedEntity">Logical name of the entity</param>
        /// <param name="service">Organization Service</param>
        /// <returns>List of views</returns>
        public static List<Entity> RetrieveViews(EntityMetadata selectedEntity, IOrganizationService service)
        {
            try
            {
                var qba = new QueryByAttribute
                {
                    EntityName = "savedquery",
                    ColumnSet = new ColumnSet(true)
                };

                qba.Attributes.Add("returnedtypecode");
                qba.Values.Add(selectedEntity.ObjectTypeCode);

                var views = service.RetrieveMultiple(qba);

                var viewsList = new List<Entity>();

                foreach (Entity entity in views.Entities)
                {
                    viewsList.Add(entity);
                }

                return viewsList;
            }
            catch (Exception error)
            {
                string errorMessage = CrmExceptionHelper.GetErrorMessage(error, false);
                throw new Exception("Error while retrieving views: " + errorMessage);
            }
        }

        /// <summary>
        /// Retrieve the list of personal views for a specific entity
        /// </summary>
        /// <param name="selectedEntity">Logical name of the entity</param>
        /// <param name="service">Organization Service</param>
        /// <returns>List of views</returns>
        internal static IEnumerable<Entity> RetrieveUserViews(EntityMetadata selectedEntity, IOrganizationService service)
        {
            try
            {

                QueryByAttribute qba = new QueryByAttribute
                {
                    EntityName = "userquery",
                    ColumnSet = new ColumnSet(true)
                };

                qba.Attributes.AddRange("returnedtypecode", "querytype");
                qba.Values.AddRange(selectedEntity.ObjectTypeCode.Value, 0);

                EntityCollection views = service.RetrieveMultiple(qba);

                List<Entity> viewsList = new List<Entity>();

                foreach (Entity entity in views.Entities)
                {
                    viewsList.Add(entity);
                }

                return viewsList;
            }
            catch (Exception error)
            {
                string errorMessage = CrmExceptionHelper.GetErrorMessage(error, false);
                throw new Exception("Error while retrieving user views: " + errorMessage);
            }
        }

    }
}
