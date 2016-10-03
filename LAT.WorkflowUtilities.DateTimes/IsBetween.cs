using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace LAT.WorkflowUtilities.DateTimes
{
    public class IsBetween : CodeActivity
    {
        [RequiredArgument]
        [Input("Starting Date")]
        public InArgument<DateTime> StartingDate { get; set; }

        [RequiredArgument]
        [Input("Date To Validate")]
        public InArgument<DateTime> DateToValidate { get; set; }

        [RequiredArgument]
        [Input("Ending Date")]
        public InArgument<DateTime> EndingDate { get; set; }

        [OutputAttribute("Between")]
        public OutArgument<bool> Between { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracer = executionContext.GetExtension<ITracingService>();

            try
            {
                DateTime startingDate = StartingDate.Get(executionContext);
                DateTime dateToValidate = DateToValidate.Get(executionContext);
                DateTime endingDate = EndingDate.Get(executionContext);

                var between = dateToValidate > startingDate && dateToValidate < endingDate;

                Between.Set(executionContext, between);
            }
            catch (Exception ex)
            {
                tracer.Trace("Exception: {0}", ex.ToString());
            }
        }
    }
}
