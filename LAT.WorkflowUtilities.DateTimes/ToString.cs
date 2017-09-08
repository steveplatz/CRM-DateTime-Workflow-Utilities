using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Globalization;
using LAT.WorkflowUtilities.DateTimes.Common;

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class ToString : CodeActivity
    {
        [RequiredArgument]
        [Input("Date To Use")]
        public InArgument<DateTime> DateToUse { get; set; }

        [RequiredArgument]
        [Input("Evaluate As User Local")]
        [Default("True")]
        public InArgument<bool> EvaluateAsUserLocal { get; set; }

        [Input("Culture")]
        [Default("en-US")]
        public InArgument<string> Culture { get; set; }

        [Input("Format")]
        [Default("g")]
        public InArgument<string> Format { get; set; }

        [OutputAttribute("Formatted Date String")]
        public OutArgument<string> FormattedDateString { get; set; }

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
                string cultureIn = Culture.Get(executionContext);
                string format = Format.Get(executionContext);

                CultureInfo culture = null;
                if (!string.IsNullOrEmpty(cultureIn))
                    culture = new CultureInfo(cultureIn);

                if (evaluateAsUserLocal)
                {
                    GetLocalTime glt = new GetLocalTime();
                    int? timeZoneCode = glt.RetrieveTimeZoneCode(service);
                    dateToUse = glt.RetrieveLocalTimeFromUtcTime(dateToUse, timeZoneCode, service);
                }

                string formattedDateString = dateToUse.ToString(format, culture);

                FormattedDateString.Set(executionContext, formattedDateString);
            }
            catch (Exception ex)
            {
                tracer.Trace("Exception: {0}", ex.ToString());
            }
        }
    }
}