using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class DateDiffHours : CodeActivity
    {
        [RequiredArgument]
        [Input("Starting Date")]
        public InArgument<DateTime> StartingDate { get; set; }

        [RequiredArgument]
        [Input("Ending Date")]
        public InArgument<DateTime> EndingDate { get; set; }

        [OutputAttribute("Hours Difference")]
        public OutArgument<int> HoursDifference { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracer = executionContext.GetExtension<ITracingService>();

            try
            {
                DateTime startingDate = StartingDate.Get(executionContext);
                DateTime endingDate = EndingDate.Get(executionContext);

                TimeSpan difference = startingDate - endingDate;

                int hoursDifference = Math.Abs(Convert.ToInt32(difference.TotalHours));

                HoursDifference.Set(executionContext, hoursDifference);
            }
            catch (Exception ex)
            {
                tracer.Trace("Exception: {0}", ex.ToString());
            }
        }
    }
}