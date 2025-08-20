using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;

namespace CrmPlugins
{
    public class LeadCreatePlugin : IPlugin
    {
        // Tamaño fijo del ciclo
        private const int CycleSize = 5;

        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = factory.CreateOrganizationService(context.UserId);
            var tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                if (!context.MessageName.Equals("Create", StringComparison.OrdinalIgnoreCase) ||
                    !context.PrimaryEntityName.Equals("lead", StringComparison.OrdinalIgnoreCase) ||
                    context.Stage != 40) // PostOperation
                {
                    return;
                }

                if (!context.OutputParameters.Contains("id") || !(context.OutputParameters["id"] is Guid leadId))
                    throw new InvalidPluginExecutionException("No se pudo obtener el id del lead.");

                var lead = service.Retrieve("lead", leadId, new ColumnSet("createdon", "ml_asesorcomercial"));
                var createdOn = (DateTime)lead["createdon"];

                var countLeads = new QueryExpression("lead")
                {
                    ColumnSet = new ColumnSet(false),
                    Criteria = new FilterExpression(LogicalOperator.And)
                };
                countLeads.Criteria.AddCondition("createdon", ConditionOperator.On, createdOn);
                var results = service.RetrieveMultiple(countLeads);
                int seq = results.Entities.Count;

                //Calculo la posicion (1..5)
                int pos = ((seq - 1) % CycleSize) + 1;

                //Asigno el asesor
                var asesor = new QueryExpression("ml_rankingdeasesores")
                {
                    ColumnSet = new ColumnSet("ml_asesorcomercial", "ml_posicionranking", "ml_activo"),
                    Criteria = new FilterExpression(LogicalOperator.And)
                };
                asesor.Criteria.AddCondition("ml_activo", ConditionOperator.Equal, true);
                asesor.Criteria.AddCondition("ml_posicionranking", ConditionOperator.Equal, pos);

                var asesores = service.RetrieveMultiple(asesor);
                if (asesores.Entities.Count != 1 || !asesores.Entities[0].Contains("ml_asesorcomercial"))
                    throw new InvalidPluginExecutionException($"No se encontró un asesor único y activo para la posición {pos}. Verificar el ranking.");

                var asesorRef = (EntityReference)asesores.Entities[0]["ml_asesorcomercial"];

                lead["ml_asesorcomercial"] = new EntityReference("systemuser", asesorRef.Id);
                service.Update(lead);

                //Guardo la secuencia en el cliente potencial
                if (!string.IsNullOrWhiteSpace("ml_numerosecuencialdeldia"))
                {
                    var updateLead = new Entity("lead", leadId);
                    updateLead["ml_numerosecuencialdeldia"] = pos;
                    service.Update(updateLead);
                }

                //Historial de asignación
                var hist = new Entity("ml_historialdeasignaciones");
                hist["ml_clientepotencial"] = new EntityReference("lead", leadId);
                hist["ml_asesor"] = new EntityReference("systemuser", asesorRef.Id);
                hist["ml_fechadeasignacion"] = DateTime.UtcNow;
                hist["ml_numerosecuencialdeldia"] = pos;
                service.Create(hist);

                tracing.Trace($"LeadCreatePlugin OK. Lead {leadId}, seq {seq}, pos {pos}, Asesor: {asesorRef.Id}");
            }
            catch (Exception ex)
            {
                tracing.Trace("LeadCreatePlugin. ERROR: {0}", ex.ToString());
                throw new InvalidPluginExecutionException("Error en la asignación de Lead: " + ex.Message, ex);
            }
        }
    }
}