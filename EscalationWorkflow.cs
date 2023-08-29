/*
Task: Escalation Workflow
Build a custom workflow for customer support cases.
If a case is not resolved within a specified time (5 dyas after created on date),
the workflow should escalate the case to a higher-level support team and send an email notification to the manager
 */

using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Crm.Sdk.Messages;
using System.Activities;

namespace CustomWorkflow
{
    public class EscalationWorkflow : CodeActivity
    {
        protected override void Execute(CodeActivityContext context)
        {
            IWorkflowContext workflowContext = context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(workflowContext.UserId);
            ITracingService tracingService = context.GetExtension<ITracingService>();

            try
            {
                tracingService.Trace("Workflow execution started.");

                Entity entityIncident = service.Retrieve("incident", workflowContext.PrimaryEntityId, new ColumnSet("title", "ticketnumber", "createdon", "ownerid"));

                if (entityIncident.Contains("ownerid") && entityIncident["ownerid"] is EntityReference ownerReference)
                {
                    // Assign the incident to a specified owner (replace GUID with actual owner ID)
                    ownerReference.Id = new Guid("67189041-c032-ee11-bdf4-000d3a0aabb1");
                    entityIncident["ownerid"] = ownerReference;

                    // Update the incident record
                    service.Update(entityIncident);

                    // Send email notification
                    Entity fromActivity = new Entity("activityparty");
                    fromActivity["partyid"] = new EntityReference("systemuser", workflowContext.InitiatingUserId);

                    Entity toActivity = new Entity("activityparty");
                    toActivity["partyid"] = new EntityReference("systemuser", Guid.Parse("67189041-c032-ee11-bdf4-000d3a0aabb1") );

                    Entity entityEmail = new Entity("email");
                    entityEmail["from"] = new Entity[] { fromActivity };
                    entityEmail["to"] = new Entity[] { toActivity };
                    entityEmail["subject"] = "Escalation for Record: " + entityIncident.GetAttributeValue<string>("title");
                    entityEmail["description"] = "This record was created on: " + entityIncident.GetAttributeValue<DateTime>("createdon") +
                                                " with Ticket Number: " + entityIncident.GetAttributeValue<string>("ticketnumber");
                    entityEmail["directioncode"] = true; // Outgoing email

                    Guid emailId = service.Create(entityEmail);

                    // Send the email
                    SendEmailRequest sendEmailRequest = new SendEmailRequest
                    {
                        EmailId = emailId,
                        TrackingToken = "",
                        IssueSend = true,
                    };
                    service.Execute(sendEmailRequest);

                    tracingService.Trace("Workflow execution completed successfully.");
                }
                else
                {
                    tracingService.Trace("Owner reference is null or invalid.");
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error: " + ex.ToString());
            }
        }
    }
}
