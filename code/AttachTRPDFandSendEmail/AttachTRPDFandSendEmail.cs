using System;xfvsvfsdvzxczxczxczxcxc
using System.Activities;
using System.ServiceModel;
//using CnCrm.WfActivities.HelperCode;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
using System.Linq;

namespace AttachTRPDFandSendEmail.AttachTRPDFandSendEmail
{
    public class Class1 : CodeActivity

    {

        //define output variable

        [Input("SourceEmail")]

        [ReferenceTarget("email")]

        public InArgument<EntityReference> SourceEmail { get; set; }

        protected override void Execute(CodeActivityContext executionContext)

        {

            // Get workflow context

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();

            //Create service factory

            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();

            // Create Organization service

            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            // Get the target entity from the context

            Entity SiteVisitID = (Entity)service.Retrieve("td_tripreportsample3", context.PrimaryEntityId, new ColumnSet(new string[] { "td_tripreportsample3id" }));

            AddAttachmentToEmailRecord(service, SiteVisitID.Id, SourceEmail.Get<EntityReference>(executionContext));

        }

        private void AddAttachmentToEmailRecord(IOrganizationService service, Guid SourceSiteVisitID, EntityReference SourceEmailID)

        {

            //create email object

            Entity _ResultEntity = service.Retrieve("email", SourceEmailID.Id, new ColumnSet(true));

            QueryExpression _QueryNotes = new QueryExpression("annotation");

            _QueryNotes.ColumnSet = new ColumnSet(new string[] { "subject", "mimetype", "filename", "documentbody" });

            _QueryNotes.Criteria = new FilterExpression();

            _QueryNotes.Criteria.FilterOperator = LogicalOperator.And;

            _QueryNotes.Criteria.AddCondition(new ConditionExpression("objectid", ConditionOperator.Equal, SourceSiteVisitID));

            EntityCollection _MimeCollection = service.RetrieveMultiple(_QueryNotes);

            if (_MimeCollection.Entities.Count > 0)

            {  //we need to fetch first attachment

                Entity _NotesAttachment = _MimeCollection.Entities.Last();

                Entity member = service.Retrieve("td_tripreportsample3", SourceSiteVisitID, new ColumnSet(true));
                String SiteVisitName = member.Attributes["td_sitevisitname"].ToString();
                String SiteVisitVersion = member.Attributes["td_reportversionnumber"].ToString();
             

             



                Entity myEntity = new Entity("annotation");

                myEntity.Id = _NotesAttachment.Id;

               

                //   EntityCollection _MimeCollection = service.RetrieveMultiple(_QueryNotes);


                //Create email attachment

                Entity _EmailAttachment = new Entity("activitymimeattachment");

                if (_NotesAttachment.Contains("subject"))

                    _EmailAttachment["subject"] = _NotesAttachment.GetAttributeValue<string>("subject");

                _EmailAttachment["objectid"] = new EntityReference("email", _ResultEntity.Id);

                _EmailAttachment["objecttypecode"] = "email";

                if (_NotesAttachment.Contains("filename"))
                    //change filel attachment name
                    _EmailAttachment["filename"] = _NotesAttachment.GetAttributeValue<string>("filename");
                   // _EmailAttachment["filename"] = newFileName;

                if (_NotesAttachment.Contains("documentbody"))

                    _EmailAttachment["body"] = _NotesAttachment.GetAttributeValue<string>("documentbody");

                if (_NotesAttachment.Contains("mimetype"))

                    _EmailAttachment["mimetype"] = _NotesAttachment.GetAttributeValue<string>("mimetype");

                service.Create(_EmailAttachment);

            }

            // Sending email

            SendEmailRequest SendEmail = new SendEmailRequest();

            SendEmail.EmailId = _ResultEntity.Id;

            SendEmail.TrackingToken = "";

            SendEmail.IssueSend = true;

            SendEmailResponse res = (SendEmailResponse)service.Execute(SendEmail);

        }

    }

}