using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using Moq;
using System;
using System.Activities;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.DateTimes.Tests
{
    [TestClass]
    public class RelativeTimeTests
    {
        #region Class Constructor
        private readonly string _namespaceClassAssembly;
        public RelativeTimeTests()
        {
            //[Namespace.class name, assembly name] for the class/assembly being tested
            //Namespace and class name can be found on the class file being tested
            //Assembly name can be found under the project properties on the Application tab
            _namespaceClassAssembly = "LAT.WorkflowUtilities.DateTimes.RelativeTime" + ", " + "LAT.WorkflowUtilities.DateTimes";
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
        public void OneSecondAgo()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 15, 0) },
                { "EndingDate", new DateTime(2015, 5, 1, 13, 15, 1) }
            };

            //Expected value
            const string expected = "one second ago";

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test
            Assert.AreEqual(expected, output["RelativeTimeString"]);
        }

        [TestMethod]
        public void SecondsAgo()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 15, 0) },
                { "EndingDate", new DateTime(2015, 5, 1, 13, 15, 10) }
            };

            //Expected value
            const string expected = "10 seconds ago";

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test
            Assert.AreEqual(expected, output["RelativeTimeString"]);
        }

        [TestMethod]
        public void MinuteAgo()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 15, 0) },
                { "EndingDate", new DateTime(2015, 5, 1, 13, 16, 2) }
            };

            //Expected value
            const string expected = "a minute ago";

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test
            Assert.AreEqual(expected, output["RelativeTimeString"]);
        }

        [TestMethod]
        public void MinutesAgo()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 15, 0) },
                { "EndingDate", new DateTime(2015, 5, 1, 13, 30, 0) }
            };

            //Expected value
            const string expected = "15 minutes ago";

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test
            Assert.AreEqual(expected, output["RelativeTimeString"]);
        }

        [TestMethod]
        public void HourAgo()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 15, 0) },
                { "EndingDate", new DateTime(2015, 5, 1, 14, 14, 0) }
            };

            //Expected value
            const string expected = "an hour ago";

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test
            Assert.AreEqual(expected, output["RelativeTimeString"]);
        }

        [TestMethod]
        public void HoursAgo()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 15, 0) },
                { "EndingDate", new DateTime(2015, 5, 1, 16, 17, 0) }
            };

            //Expected value
            const string expected = "3 hours ago";

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test
            Assert.AreEqual(expected, output["RelativeTimeString"]);
        }

        [TestMethod]
        public void Yesterday()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 15, 0) },
                { "EndingDate", new DateTime(2015, 5, 2, 14, 14, 0) }
            };

            //Expected value
            const string expected = "yesterday";

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test
            Assert.AreEqual(expected, output["RelativeTimeString"]);
        }

        [TestMethod]
        public void DaysAgo()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 15, 0) },
                { "EndingDate", new DateTime(2015, 5, 4, 14, 14, 0) }
            };

            //Expected value
            const string expected = "3 days ago";

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test
            Assert.AreEqual(expected, output["RelativeTimeString"]);
        }

        [TestMethod]
        public void OneMonthAgo()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 15, 0) },
                { "EndingDate", new DateTime(2015, 6, 1, 14, 14, 0) }
            };

            //Expected value
            const string expected = "one month ago";

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test
            Assert.AreEqual(expected, output["RelativeTimeString"]);
        }

        [TestMethod]
        public void MonthsAgo()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 15, 0) },
                { "EndingDate", new DateTime(2015, 7, 1, 14, 14, 0) }
            };

            //Expected value
            const string expected = "2 months ago";

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test
            Assert.AreEqual(expected, output["RelativeTimeString"]);
        }

        [TestMethod]
        public void OneYearAgo()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 15, 0) },
                { "EndingDate", new DateTime(2016, 5, 1, 14, 14, 0) }
            };

            //Expected value
            const string expected = "one year ago";

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test
            Assert.AreEqual(expected, output["RelativeTimeString"]);
        }

        [TestMethod]
        public void YearsAgo()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 15, 0) },
                { "EndingDate", new DateTime(2017, 5, 1, 14, 14, 0) }
            };

            //Expected value
            const string expected = "2 years ago";

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test
            Assert.AreEqual(expected, output["RelativeTimeString"]);
        }

        [TestMethod]
        public void Future()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "StartingDate", new DateTime(2015, 5, 1, 13, 30, 0) },
                { "EndingDate", new DateTime(2015, 5, 1, 13, 15, 0) }
            };

            //Expected value
            const string expected = "in the future";

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test
            Assert.AreEqual(expected, output["RelativeTimeString"]);
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
