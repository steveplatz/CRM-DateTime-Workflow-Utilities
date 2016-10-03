using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace LAT.WorkflowUtilities.DateTimes
{
    public class ToDateTime : CodeActivity
    {
        [RequiredArgument]
        [Input("Text To Convert")]
        public InArgument<string> TextToConvert { get; set; }

        [OutputAttribute("Converted Date")]
        public OutArgument<DateTime> ConvertedDate { get; set; }

        [OutputAttribute("Is Valid Date")]
        public OutArgument<bool> IsValid { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracer = executionContext.GetExtension<ITracingService>();

            try
            {
                string textToConvert = TextToConvert.Get(executionContext);

                DateTime convertedDate;
                bool isValid = DateTime.TryParse(textToConvert, out convertedDate);

                ConvertedDate.Set(executionContext, convertedDate);
                IsValid.Set(executionContext, isValid);
            }
            catch (Exception ex)
            {
                tracer.Trace("Exception: {0}", ex.ToString());
            }
        }
    }
}
