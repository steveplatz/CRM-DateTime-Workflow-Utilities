using LAT.WorkflowUtilities.DateTimes.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace LAT.WorkflowUtilities.DateTimes
{
    public sealed class IsBusinessDay : CodeActivity
    {
        [RequiredArgument]
        [Input("Date To Check")]
        public InArgument<DateTime> DateToCheck { get; set; }

        [Input("Holiday/Closure Calendar")]
        [ReferenceTarget("calendar")]
        public InArgument<EntityReference> HolidayClosureCalendar { get; set; }

        [RequiredArgument]
        [Input("Evaluate As User Local")]
        [Default("True")]
        public InArgument<bool> EvaluateAsUserLocal { get; set; }

        [OutputAttribute("Valid Business Day")]
        public OutArgument<bool> ValidBusinessDay { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracer = executionContext.GetExtension<ITracingService>();
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                DateTime dateToCheck = DateToCheck.Get(executionContext);
                bool evaluateAsUserLocal = EvaluateAsUserLocal.Get(executionContext);

                if (evaluateAsUserLocal)
                {
                    GetLocalTime glt = new GetLocalTime();
                    int? timeZoneCode = glt.RetrieveTimeZoneCode(service);
                    dateToCheck = glt.RetrieveLocalTimeFromUtcTime(dateToCheck, timeZoneCode, service);
                }

                EntityReference holidaySchedule = HolidayClosureCalendar.Get(executionContext);

                bool validBusinessDay = dateToCheck.DayOfWeek != DayOfWeek.Saturday || dateToCheck.DayOfWeek == DayOfWeek.Sunday;

                if (!validBusinessDay)
                {
                    ValidBusinessDay.Set(executionContext, false);
                    return;
                }

                if (holidaySchedule != null)
                {
                    Entity calendar = service.Retrieve("calendar", holidaySchedule.Id, new ColumnSet(true));
                    if (calendar == null) return;

                    EntityCollection calendarRules = calendar.GetAttributeValue<EntityCollection>("calendarrules");
                    foreach (Entity calendarRule in calendarRules.Entities)
                    {
                        //Date is not stored as UTC
                        DateTime startTime = calendarRule.GetAttributeValue<DateTime>("starttime");

                        //Not same date
                        if (!startTime.Date.Equals(dateToCheck.Date))
                            continue;

                        //Not full day event
                        if (startTime.Subtract(startTime.TimeOfDay) != startTime || calendarRule.GetAttributeValue<int>("duration") != 1440)
                            continue;

                        ValidBusinessDay.Set(executionContext, false);
                        return;
                    }
                }

                ValidBusinessDay.Set(executionContext, true);
            }
            catch (Exception ex)
            {
                tracer.Trace("Exception: {0}", ex.ToString());
            }
        }
    }
}