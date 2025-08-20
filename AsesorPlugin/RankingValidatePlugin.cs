using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;

namespace AsesorPlugin
{
    public class RankingValidatePlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = factory.CreateOrganizationService(context.UserId);

            if ((context.MessageName == "Create" || context.MessageName == "Update") &&
                context.PrimaryEntityName == "ml_rankingdeasesores")
            {
                var rankingRecord = (Entity)context.InputParameters["Target"];
                int pos = -1;
                var posOption = rankingRecord.GetAttributeValue<OptionSetValue>("ml_posicionranking");
                if (posOption != null)
                {
                    pos = posOption.Value;
                }
                else //Cuando cambia estado no se envía la posición
                { 
                    if (context.MessageName == "Update" && context.PreEntityImages.Contains("PreImage"))
                    {
                        var preImage = context.PreEntityImages["PreImage"];
                        var prePosOption = preImage.GetAttributeValue<OptionSetValue>("ml_posicionranking");
                        if (prePosOption != null)
                        {
                            pos = prePosOption.Value;
                        }
                    }
                }

                if (pos < 1 || pos > 5)
                    throw new InvalidPluginExecutionException("La posición de ranking debe ser un valor entre 1 y 5.");

                var query = new QueryExpression("ml_rankingdeasesores")
                {
                    ColumnSet = new ColumnSet("ml_posicionranking")
                };
                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); // activos

                if (rankingRecord.Id != Guid.Empty) //Update
                {
                    query.Criteria.AddCondition("ml_rankingdeasesoresid", ConditionOperator.NotEqual, rankingRecord.Id.ToString());
                }

                var existingRanking = service.RetrieveMultiple(query);

                if (existingRanking.Entities.Count >= 5)
                    throw new InvalidPluginExecutionException("No se pueden tener más de 5 asesores activos en el ranking.");

                // Verificar duplicados de posición
                bool duplicatePosition = existingRanking.Entities
                    .Any(e => e.Contains("ml_posicionranking") &&
                              e.GetAttributeValue<OptionSetValue>("ml_posicionranking").Value == pos);

                if (duplicatePosition)
                    throw new InvalidPluginExecutionException($"Ya existe un asesor con la posición {pos} en el ranking.");
            }
        }
    }

}
