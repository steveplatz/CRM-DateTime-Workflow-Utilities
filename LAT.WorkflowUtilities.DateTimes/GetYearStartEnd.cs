using LAT.WorkflowUtilities.DateTimes.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace LAT.WorkflowUtilities.DateTimes
{
    public class GetYearStartEnd : CodeActivity
    {
        [RequiredArgument]
        [Input("Date To Use")]
        public InArgument<DateTime> DateToUse { get; set; }

        [RequiredArgument]
        [Input("Evaluate As User Local")]
        [Default("True")]
        public InArgument<bool> EvaluateAsUserLocal { get; set; }

        [OutputAttribute("Year Start Date")]
        public OutArgument<DateTime> YearStartDate { get; set; }

        [OutputAttribute("Year End Date")]
        public OutArgument<DateTime> YearEndDate { get; set; }

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

                DateTime yearStartDate = new DateTime(dateToUse.Year, 1, 1, 0, 0, 0);
                DateTime yearEndDate = yearStartDate.AddYears(1).AddDays(-1).AddHours(23).AddMinutes(59).AddSeconds(59).AddMilliseconds(999);

                YearStartDate.Set(executionContext, yearStartDate);
                YearEndDate.Set(executionContext, yearEndDate);
            }
            catch (Exception ex)
            {
                tracer.Trace("Exception: {0}", ex.ToString());
            }
        }
    }
}
