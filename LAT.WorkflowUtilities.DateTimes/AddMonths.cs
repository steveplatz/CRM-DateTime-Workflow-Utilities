using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class AddMonths : CodeActivity
    {
        [RequiredArgument]
        [Input("Original Date")]
        public InArgument<DateTime> OriginalDate { get; set; }

        [RequiredArgument]
        [Input("Months To Add")]
        public InArgument<int> MonthsToAdd { get; set; }

        [OutputAttribute("Updated Date")]
        public OutArgument<DateTime> UpdatedDate { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracer = executionContext.GetExtension<ITracingService>();

            try
            {
                DateTime originalDate = OriginalDate.Get(executionContext);
                int monthsToAdd = MonthsToAdd.Get(executionContext);

                DateTime updatedDate = originalDate.AddMonths(monthsToAdd);

                UpdatedDate.Set(executionContext, updatedDate);
            }
            catch (Exception ex)
            {
                tracer.Trace("Exception: {0}", ex.ToString());
            }
        }
    }
}