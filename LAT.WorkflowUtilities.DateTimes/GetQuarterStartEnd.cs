using LAT.WorkflowUtilities.DateTimes.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace LAT.WorkflowUtilities.DateTimes
{
    public class GetQuarterStartEnd : CodeActivity
    {
        [RequiredArgument]
        [Input("Date To Use")]
        public InArgument<DateTime> DateToUse { get; set; }

        [RequiredArgument]
        [Input("Evaluate As User Local")]
        [Default("True")]
        public InArgument<bool> EvaluateAsUserLocal { get; set; }

        [OutputAttribute("Quarter Start Date")]
        public OutArgument<DateTime> QuarterStartDate { get; set; }

        [OutputAttribute("Quarter End Date")]
        public OutArgument<DateTime> QuarterEndDate { get; set; }

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

                int quarterNumber = (dateToUse.Month - 1) / 3 + 1;
                DateTime quarterStartDate = new DateTime(dateToUse.Year, (quarterNumber - 1) * 3 + 1, 1, 0, 0, 0);
                DateTime quarterEndDate = quarterStartDate.AddMonths(3).AddDays(-1).AddHours(23).AddMinutes(59).AddSeconds(59).AddMilliseconds(999);

                QuarterStartDate.Set(executionContext, quarterStartDate);
                QuarterEndDate.Set(executionContext, quarterEndDate);
            }
            catch (Exception ex)
            {
                tracer.Trace("Exception: {0}", ex.ToString());
            }
        }
    }
}
