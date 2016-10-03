using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class DateDiffMinutes : CodeActivity
    {
        [RequiredArgument]
        [Input("Starting Date")]
        public InArgument<DateTime> StartingDate { get; set; }

        [RequiredArgument]
        [Input("Ending Date")]
        public InArgument<DateTime> EndingDate { get; set; }

        [OutputAttribute("Minutes Difference")]
        public OutArgument<int> MinutesDifference { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracer = executionContext.GetExtension<ITracingService>();

            try
            {
                DateTime startingDate = StartingDate.Get(executionContext);
                DateTime endingDate = EndingDate.Get(executionContext);

                startingDate = new DateTime(startingDate.Year, startingDate.Month, startingDate.Day, startingDate.Hour,
                    startingDate.Minute, 0, startingDate.Kind);

                endingDate = new DateTime(endingDate.Year, endingDate.Month, endingDate.Day, endingDate.Hour,
                    endingDate.Minute, 0, endingDate.Kind);

                TimeSpan difference = startingDate - endingDate;

                int minutesDifference = Math.Abs(Convert.ToInt32(difference.TotalMinutes));

                MinutesDifference.Set(executionContext, minutesDifference);
            }
            catch (Exception ex)
            {
                tracer.Trace("Exception: {0}", ex.ToString());
            }
        }
    }
}