using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using Moq;
using System;
using System.Activities;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.DateTimes.Tests
{
    [TestClass]
    public class AddBusinessDaysTests
    {
        #region Class Constructor
        private readonly string _namespaceClassAssembly;
        public AddBusinessDaysTests()
        {
            //[Namespace.class name, assembly name] for the class/assembly being tested
            //Namespace and class name can be found on the class file being tested
            //Assembly name can be found under the project properties on the Application tab
            _namespaceClassAssembly = "LAT.WorkflowUtilities.DateTimes.AddBusinessDays" + ", " + "LAT.WorkflowUtilities.DateTimes";
        }
        #endregion
        #region Test Initialization and Cleanup
        // Use ClassInitialize to run code before running the first test in the class
        [ClassInitialize()]
        public static void ClassInitialize(TestContext testContext) { }

        // Use ClassCleanup to run code after all tests in a class have run
        [ClassCleanup()]
        public static void ClassCleanup() { }

        // Use TestInitialize to run code before running each test 
        [TestInitialize()]
        public void TestMethodInitialize() { }

        // Use TestCleanup to run code after each test has run
        [TestCleanup()]
        public void TestMethodCleanup() { }
        #endregion

        [TestMethod]
        public void AddOneNoWeekend()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "OriginalDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "BusinessDaysToAdd", 1 },
                { "HolidayClosureCalendar", null }
            };

            //Expected value
            DateTime expected = new DateTime(2014, 7, 4, 8, 48, 0, 0);

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test
            Assert.AreEqual(expected, output["UpdatedDate"]);
        }

        [TestMethod]
        public void AddThreeWeekend()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "OriginalDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "BusinessDaysToAdd", 3 },
                { "HolidayClosureCalendar", null }
            };

            //Expected value
            DateTime expected = new DateTime(2014, 7, 8, 8, 48, 0, 0);

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test
            Assert.AreEqual(expected, output["UpdatedDate"]);
        }

        [TestMethod]
        public void AddMinusOneNoWeekend()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "OriginalDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "BusinessDaysToAdd", -1 },
                { "HolidayClosureCalendar", null }
            };

            //Expected value
            DateTime expected = new DateTime(2014, 7, 2, 8, 48, 0, 0);

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test
            Assert.AreEqual(expected, output["UpdatedDate"]);
        }

        [TestMethod]
        public void AddMinusFiveWeekend()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "OriginalDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "BusinessDaysToAdd", -5 },
                { "HolidayClosureCalendar", null }
            };

            //Expected value
            DateTime expected = new DateTime(2014, 6, 26, 8, 48, 0, 0);

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test
            Assert.AreEqual(expected, output["UpdatedDate"]);
        }

        [TestMethod]
        public void AddOneNonFullDayHoliday()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "OriginalDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "BusinessDaysToAdd", 1 },
                { "HolidayClosureCalendar", new EntityReference{LogicalName = "calendar", Id = new Guid("e9717a91-ba0a-e411-b681-6c3be5a8ad70")}}
            };

            //Expected value
            DateTime expected = new DateTime(2014, 7, 4, 8, 48, 0, 0);

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, AddOneNonFullDayHolidaySetup);

            //Test
            Assert.AreEqual(expected, output["UpdatedDate"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> AddOneNonFullDayHolidaySetup(Mock<IOrganizationService> serviceMock)
        {
            EntityCollection calendarRules = new EntityCollection();
            Entity calendarRule = new Entity
            {
                LogicalName = "calendarRule",
                Id = Guid.NewGuid(),
                Attributes = new AttributeCollection()
            };
            calendarRule.Attributes.Add("name", "4th of July");
            calendarRule.Attributes.Add("starttime", new DateTime(2014, 7, 4, 11, 0, 0));
            calendarRule.Attributes.Add("duration", 30);
            calendarRules.Entities.Add(calendarRule);

            Entity holidayCalendar = new Entity
            {
                Id = new Guid("e9717a91-ba0a-e411-b681-6c3be5a8ad70"),
                LogicalName = "calendar",
                Attributes = new AttributeCollection()
            };
            holidayCalendar.Attributes.Add("name", "Main Holiday Schedule");
            holidayCalendar.Attributes.Add("calendarrules", calendarRules);

            serviceMock.Setup(t =>
                t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
                .ReturnsInOrder(holidayCalendar);

            return serviceMock;
        }

        [TestMethod]
        public void AddOneNonFullDayHolidayPlusWeekend()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "OriginalDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "BusinessDaysToAdd", 4 },
                { "HolidayClosureCalendar", new EntityReference{LogicalName = "calendar", Id = new Guid("e9717a91-ba0a-e411-b681-6c3be5a8ad70")}}
            };

            //Expected value
            DateTime expected = new DateTime(2014, 7, 9, 8, 48, 0, 0);

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, AddOneNonFullDayHolidayPlusWeekendSetup);

            //Test
            Assert.AreEqual(expected, output["UpdatedDate"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> AddOneNonFullDayHolidayPlusWeekendSetup(Mock<IOrganizationService> serviceMock)
        {
            EntityCollection calendarRules = new EntityCollection();
            Entity calendarRule = new Entity
            {
                LogicalName = "calendarRule",
                Id = Guid.NewGuid(),
                Attributes = new AttributeCollection()
            };
            calendarRule.Attributes.Add("name", "4th of July");
            calendarRule.Attributes.Add("starttime", new DateTime(2014, 7, 4, 11, 0, 0));
            calendarRule.Attributes.Add("duration", 30);
            calendarRules.Entities.Add(calendarRule);

            Entity holidayCalendar = new Entity
            {
                Id = new Guid("e9717a91-ba0a-e411-b681-6c3be5a8ad70"),
                LogicalName = "calendar",
                Attributes = new AttributeCollection()
            };
            holidayCalendar.Attributes.Add("name", "Main Holiday Schedule");
            holidayCalendar.Attributes.Add("calendarrules", calendarRules);

            serviceMock.Setup(t =>
                t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
                .ReturnsInOrder(holidayCalendar);

            return serviceMock;
        }

        [TestMethod]
        public void AddOneHolidayPlusWeekend()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
	        {
		        { "OriginalDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
		        { "BusinessDaysToAdd", 1 },
		        { "HolidayClosureCalendar", new EntityReference{LogicalName = "calendar", Id = new Guid("e9717a91-ba0a-e411-b681-6c3be5a8ad70")}}
	        };

            //Expected value
            DateTime expected = new DateTime(2014, 7, 7, 8, 48, 0, 0);

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, AddOneHolidayPlusWeekendSetup);

            //Test
            Assert.AreEqual(expected, output["UpdatedDate"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> AddOneHolidayPlusWeekendSetup(Mock<IOrganizationService> serviceMock)
        {
            EntityCollection calendarRules = new EntityCollection();
            Entity calendarRule = new Entity
            {
                LogicalName = "calendarRule",
                Id = Guid.NewGuid(),
                Attributes = new AttributeCollection()
            };
            calendarRule.Attributes.Add("name", "4th of July");
            calendarRule.Attributes.Add("starttime", new DateTime(2014, 7, 4, 0, 0, 0));
            calendarRule.Attributes.Add("duration", 1440);
            calendarRules.Entities.Add(calendarRule);

            Entity holidayCalendar = new Entity
            {
                Id = new Guid("e9717a91-ba0a-e411-b681-6c3be5a8ad70"),
                LogicalName = "calendar",
                Attributes = new AttributeCollection()
            };
            holidayCalendar.Attributes.Add("name", "Main Holiday Schedule");
            holidayCalendar.Attributes.Add("calendarrules", calendarRules);

            serviceMock.Setup(t =>
                t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
                .ReturnsInOrder(holidayCalendar);

            return serviceMock;
        }

        [TestMethod]
        public void AddOneNonFullDayBusinessClosure()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "OriginalDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "BusinessDaysToAdd", 1 },
                { "HolidayClosureCalendar", new EntityReference{LogicalName = "calendar", Id = new Guid("b01748c5-d0ba-e311-9ec9-6c3be5a8a0c8")}}
            };

            //Expected value
            DateTime expected = new DateTime(2014, 7, 4, 8, 48, 0, 0);

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, AddOneNonFullDayBusinessClosureSetup);

            //Test
            Assert.AreEqual(expected, output["UpdatedDate"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> AddOneNonFullDayBusinessClosureSetup(Mock<IOrganizationService> serviceMock)
        {
            EntityCollection calendarRules = new EntityCollection();
            Entity calendarRule = new Entity
            {
                LogicalName = "calendarRule",
                Id = Guid.NewGuid(),
                Attributes = new AttributeCollection()
            };
            calendarRule.Attributes.Add("name", "Some Half Day");
            calendarRule.Attributes.Add("starttime", new DateTime(2014, 7, 4, 8, 0, 0));
            calendarRule.Attributes.Add("duration", 240);
            calendarRules.Entities.Add(calendarRule);

            Entity holidayCalendar = new Entity
            {
                Id = new Guid("b01748c5-d0ba-e311-9ec9-6c3be5a8a0c8"),
                LogicalName = "calendar",
                Attributes = new AttributeCollection()
            };
            holidayCalendar.Attributes.Add("name", "Business Closure Calendar");
            holidayCalendar.Attributes.Add("calendarrules", calendarRules);

            serviceMock.Setup(t =>
                t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
                .ReturnsInOrder(holidayCalendar);

            return serviceMock;
        }

        public void AddOneBusinessClosurePlusWeekend()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "OriginalDate", new DateTime(2014, 7, 3, 8, 48, 0, 0)},
                { "BusinessDaysToAdd", 1 },
                { "HolidayClosureCalendar", new EntityReference{LogicalName = "calendar", Id = new Guid("b01748c5-d0ba-e311-9ec9-6c3be5a8a0c8")}}
            };

            //Expected value
            DateTime expected = new DateTime(2014, 7, 7, 8, 48, 0, 0);

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, AddOneBusinessClosurePlusWeekendSetup);

            //Test
            Assert.AreEqual(expected, output["UpdatedDate"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> AddOneBusinessClosurePlusWeekendSetup(Mock<IOrganizationService> serviceMock)
        {
            EntityCollection calendarRules = new EntityCollection();
            Entity calendarRule = new Entity
            {
                LogicalName = "calendarRule",
                Id = Guid.NewGuid(),
                Attributes = new AttributeCollection()
            };
            calendarRule.Attributes.Add("name", "4th of July");
            calendarRule.Attributes.Add("starttime", new DateTime(2014, 7, 4, 0, 0, 0));
            calendarRule.Attributes.Add("duration", 1440);
            calendarRules.Entities.Add(calendarRule);

            Entity holidayCalendar = new Entity
            {
                Id = new Guid("b01748c5-d0ba-e311-9ec9-6c3be5a8a0c8"),
                LogicalName = "calendar",
                Attributes = new AttributeCollection()
            };
            holidayCalendar.Attributes.Add("name", "Business Closure Calendar");
            holidayCalendar.Attributes.Add("calendarrules", calendarRules);

            serviceMock.Setup(t =>
                t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
                .ReturnsInOrder(holidayCalendar);

            return serviceMock;
        }

        /// <summary>
        /// Invokes the workflow.
        /// </summary>
        /// <param name="name">Namespace.Class, Assembly</param>
        /// <param name="target">The target entity</param>
        /// <param name="inputs">The workflow input parameters</param>
        /// <param name="configuredServiceMock">The function to configure the Organization Service</param>
        /// <returns>The workflow output parameters</returns>
        private static IDictionary<string, object> InvokeWorkflow(string name, ref Entity target, Dictionary<string, object> inputs,
            Func<Mock<IOrganizationService>, Mock<IOrganizationService>> configuredServiceMock)
        {
            var testClass = Activator.CreateInstance(Type.GetType(name)) as CodeActivity; ;

            var serviceMock = new Mock<IOrganizationService>();
            var factoryMock = new Mock<IOrganizationServiceFactory>();
            var tracingServiceMock = new Mock<ITracingService>();
            var workflowContextMock = new Mock<IWorkflowContext>();

            //Apply configured Organization Service Mock
            if (configuredServiceMock != null)
                serviceMock = configuredServiceMock(serviceMock);

            IOrganizationService service = serviceMock.Object;

            //Mock workflow Context
            var workflowUserId = Guid.NewGuid();
            var workflowCorrelationId = Guid.NewGuid();
            var workflowInitiatingUserId = Guid.NewGuid();

            //Workflow Context Mock
            workflowContextMock.Setup(t => t.InitiatingUserId).Returns(workflowInitiatingUserId);
            workflowContextMock.Setup(t => t.CorrelationId).Returns(workflowCorrelationId);
            workflowContextMock.Setup(t => t.UserId).Returns(workflowUserId);
            var workflowContext = workflowContextMock.Object;

            //Organization Service Factory Mock
            factoryMock.Setup(t => t.CreateOrganizationService(It.IsAny<Guid>())).Returns(service);
            var factory = factoryMock.Object;

            //Tracing Service - Content written appears in output
            tracingServiceMock.Setup(t => t.Trace(It.IsAny<string>(), It.IsAny<object[]>())).Callback<string, object[]>(MoqExtensions.WriteTrace);
            var tracingService = tracingServiceMock.Object;

            //Parameter Collection
            ParameterCollection inputParameters = new ParameterCollection { { "Target", target } };
            workflowContextMock.Setup(t => t.InputParameters).Returns(inputParameters);

            //Workflow Invoker
            var invoker = new WorkflowInvoker(testClass);
            invoker.Extensions.Add(() => tracingService);
            invoker.Extensions.Add(() => workflowContext);
            invoker.Extensions.Add(() => factory);

            return invoker.Invoke(inputs);
        }
    }
}
