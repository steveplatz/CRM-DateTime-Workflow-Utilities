using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Globalization;

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class ToUTCString : CodeActivity
    {
        [RequiredArgument]
        [Input("Date To Format")]
        public InArgument<DateTime> DateToFormat { get; set; }

        [Input("Culture")]
        [Default("en-US")]
        public InArgument<string> Culture { get; set; }

        [OutputAttribute("Formatted Date String")]
        public OutArgument<string> FormattedDateString { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracer = executionContext.GetExtension<ITracingService>();

            try
            {
                DateTime originalDate = DateToFormat.Get(executionContext);
                string cultureIn = Culture.Get(executionContext);

                CultureInfo culture = null;
                if (!string.IsNullOrEmpty(cultureIn))
                    culture = new CultureInfo(cultureIn);

                string formattedDateString = originalDate.ToUniversalTime().ToString(culture);

                FormattedDateString.Set(executionContext, formattedDateString);
            }
            catch (Exception ex)
            {
                tracer.Trace("Exception: {0}", ex.ToString());
            }
        }
    }
}