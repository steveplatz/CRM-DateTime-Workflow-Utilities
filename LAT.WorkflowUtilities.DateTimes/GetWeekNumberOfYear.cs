using LAT.WorkflowUtilities.DateTimes.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Globalization;

namespace LAT.WorkflowUtilities.DateTimes
{
    public class GetWeekNumberOfYear : CodeActivity
    {
        [RequiredArgument]
        [Input("Date To Use")]
        public InArgument<DateTime> DateToUse { get; set; }

        [RequiredArgument]
        [Input("Evaluate As User Local")]
        [Default("True")]
        public InArgument<bool> EvaluateAsUserLocal { get; set; }

        [OutputAttribute("Week Number Of Year")]
        public OutArgument<int> WeekNumberOfYear { get; set; }

        //http://blogs.msdn.com/b/shawnste/archive/2006/01/24/iso-8601-week-of-year-format-in-microsoft-net.aspx

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

                DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
                Calendar cal = dfi.Calendar;

                int weekNumberOfYear = cal.GetWeekOfYear(dateToUse, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);

                WeekNumberOfYear.Set(executionContext, weekNumberOfYear);
            }
            catch (Exception ex)
            {
                tracer.Trace("Exception: {0}", ex.ToString());
            }
        }
    }
}
