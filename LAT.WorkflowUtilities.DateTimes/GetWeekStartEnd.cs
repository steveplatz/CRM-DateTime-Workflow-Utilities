using LAT.WorkflowUtilities.DateTimes.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace LAT.WorkflowUtilities.DateTimes
{
    public class GetWeekStartEnd : CodeActivity
    {
        [RequiredArgument]
        [Input("Date To Use")]
        public InArgument<DateTime> DateToUse { get; set; }

        [RequiredArgument]
        [Input("Evaluate As User Local")]
        [Default("True")]
        public InArgument<bool> EvaluateAsUserLocal { get; set; }

        [OutputAttribute("Week Start Date")]
        public OutArgument<DateTime> WeekStartDate { get; set; }

        [OutputAttribute("Week End Date")]
        public OutArgument<DateTime> WeekEndDate { get; set; }

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

                int diff = dateToUse.DayOfWeek - DayOfWeek.Sunday;
                if (diff < 0)
                    diff += 7;

                DateTime weekStartDate = dateToUse.AddDays(-1 * diff).Date;
                weekStartDate = new DateTime(weekStartDate.Year, weekStartDate.Month, weekStartDate.Day, 0, 0, 0);
                DateTime weekEndDate = weekStartDate.AddDays(6).AddHours(23).AddMinutes(59).AddSeconds(59).AddMilliseconds(999);

                WeekStartDate.Set(executionContext, weekStartDate);
                WeekEndDate.Set(executionContext, weekEndDate);
            }
            catch (Exception ex)
            {
                tracer.Trace("Exception: {0}", ex.ToString());
            }
        }
    }
}
