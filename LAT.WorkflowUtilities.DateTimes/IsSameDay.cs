using LAT.WorkflowUtilities.DateTimes.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace LAT.WorkflowUtilities.DateTimes
{
    public class IsSameDay : CodeActivity
    {
        [RequiredArgument]
        [Input("First Date")]
        public InArgument<DateTime> FirstDate { get; set; }

        [RequiredArgument]
        [Input("Second Date")]
        public InArgument<DateTime> SecondDate { get; set; }

        [RequiredArgument]
        [Input("Evaluate As User Local")]
        [Default("True")]
        public InArgument<bool> EvaluateAsUserLocal { get; set; }

        [OutputAttribute("Same Day")]
        public OutArgument<bool> SameDay { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracer = executionContext.GetExtension<ITracingService>();
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
           
            try
            {
                DateTime firstDate = FirstDate.Get(executionContext);
                DateTime secondDate = SecondDate.Get(executionContext);
                bool evaluateAsUserLocal = EvaluateAsUserLocal.Get(executionContext);

                if (evaluateAsUserLocal)
                {
                    GetLocalTime glt = new GetLocalTime();
                    int? timeZoneCode = glt.RetrieveTimeZoneCode(service);
                    firstDate = glt.RetrieveLocalTimeFromUtcTime(firstDate, timeZoneCode, service);
                    secondDate = glt.RetrieveLocalTimeFromUtcTime(secondDate, timeZoneCode, service);
                }

                bool sameDay = firstDate.Date == secondDate.Date;

                SameDay.Set(executionContext, sameDay);
            }
            catch (Exception ex)
            {
                tracer.Trace("Exception: {0}", ex.ToString());
            }
        }
    }
}
