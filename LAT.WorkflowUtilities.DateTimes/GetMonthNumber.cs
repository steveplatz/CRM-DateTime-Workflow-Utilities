using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using LAT.WorkflowUtilities.DateTimes.Common;

namespace LAT.WorkflowUtilities.DateTimes
{
    public class GetMonthNumber : CodeActivity
    {
        [RequiredArgument]
        [Input("Date To Use")]
        public InArgument<DateTime> DateToUse { get; set; }

        [RequiredArgument]
        [Input("Evaluate As User Local")]
        [Default("True")]
        public InArgument<bool> EvaluateAsUserLocal { get; set; }

        [OutputAttribute("Month Number")]
        public OutArgument<int> MonthNumber { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracer = executionContext.GetExtension<ITracingService>();
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                DateTime dateToUse = DateToUse.Get(executionContext);
                bool evaluateAsUserLocal = EvaluateAsUserLocal.Get(executionContext);

                if (evaluateAsUserLocal)
                {
                    GetLocalTime glt = new GetLocalTime();
                    int? timeZoneCode = glt.RetrieveTimeZoneCode(service);
                    dateToUse = glt.RetrieveLocalTimeFromUtcTime(dateToUse, timeZoneCode, service);
                }

                int monthNumber = dateToUse.Month;

                MonthNumber.Set(executionContext, monthNumber);
            }
            catch (Exception ex)
            {
                tracer.Trace("Exception: {0}", ex.ToString());
            }
        }
    }
}
