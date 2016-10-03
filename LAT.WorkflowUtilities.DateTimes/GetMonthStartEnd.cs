using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using LAT.WorkflowUtilities.DateTimes.Common;

namespace LAT.WorkflowUtilities.DateTimes
{
    public class GetMonthStartEnd : CodeActivity
    {
        [RequiredArgument]
        [Input("Date To Use")]
        public InArgument<DateTime> DateToUse { get; set; }

        [RequiredArgument]
        [Input("Evaluate As User Local")]
        [Default("True")]
        public InArgument<bool> EvaluateAsUserLocal { get; set; }

        [OutputAttribute("Month Start Date")]
        public OutArgument<DateTime> MonthStartDate { get; set; }

        [OutputAttribute("Month End Date")]
        public OutArgument<DateTime> MonthEndDate { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracer = executionContext.GetExtension<ITracingService>();
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                DateTime dateToUse = DateToUse.Get(executionContext); ;
                bool evaluateAsUserLocal = EvaluateAsUserLocal.Get(executionContext);

                if (evaluateAsUserLocal)
                {
                    GetLocalTime glt = new GetLocalTime();
                    int? timeZoneCode = glt.RetrieveTimeZoneCode(service);
                    dateToUse = glt.RetrieveLocalTimeFromUtcTime(dateToUse, timeZoneCode, service);
                }

                DateTime monthStartDate = new DateTime(dateToUse.Year, dateToUse.Month, 1, 0, 0, 0);
                DateTime monthEndDate = monthStartDate.AddMonths(1).AddDays(-1).AddHours(23).AddMinutes(59).AddSeconds(59).AddMilliseconds(999);

                MonthStartDate.Set(executionContext, monthStartDate);
                MonthEndDate.Set(executionContext, monthEndDate);
            }
            catch (Exception ex)
            {
                tracer.Trace("Exception: {0}", ex.ToString());
            }
        }
    }
}
