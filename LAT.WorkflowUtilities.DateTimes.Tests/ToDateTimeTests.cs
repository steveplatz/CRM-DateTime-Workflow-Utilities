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
    public class ToDateTimeTests
    {
        #region Class Constructor
        private readonly string _namespaceClassAssembly;
        public ToDateTimeTests()
        {
            //[Namespace.class name, assembly name] for the class/assembly being tested
            //Namespace and class name can be found on the class file being tested
            //Assembly name can be found under the project properties on the Application tab
            _namespaceClassAssembly = "LAT.WorkflowUtilities.DateTimes.ToDateTime" + ", " + "LAT.WorkflowUtilities.DateTimes";
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
        public void ValidDateOnly1()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "TextToConvert", "5/1/2015" }
            };

            //Expected value(s)
            DateTime expected1 = new DateTime(2015, 5, 1, 0, 0, 0);
            const bool expected2 = true;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test(s)
            Assert.AreEqual(expected1, output["ConvertedDate"]);
            Assert.AreEqual(expected2, output["IsValid"]);
        }

        [TestMethod]
        public void ValidDateOnly2()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "TextToConvert", "05/01/2015" }
            };

            //Expected value(s)
            DateTime expected1 = new DateTime(2015, 5, 1, 0, 0, 0);
            const bool expected2 = true;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test(s)
            Assert.AreEqual(expected1, output["ConvertedDate"]);
            Assert.AreEqual(expected2, output["IsValid"]);
        }

        [TestMethod]
        public void ValidDateOnly3()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "TextToConvert", "5-1-2015" }
            };

            //Expected value(s)
            DateTime expected1 = new DateTime(2015, 5, 1, 0, 0, 0);
            const bool expected2 = true;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test(s)
            Assert.AreEqual(expected1, output["ConvertedDate"]);
            Assert.AreEqual(expected2, output["IsValid"]);
        }

        [TestMethod]
        public void ValidDateTime1()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "TextToConvert", "5/1/2015 15:00:00" }
            };

            //Expected value(s)
            DateTime expected1 = new DateTime(2015, 5, 1, 15, 0, 0);
            const bool expected2 = true;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test(s)
            Assert.AreEqual(expected1, output["ConvertedDate"]);
            Assert.AreEqual(expected2, output["IsValid"]);
        }

        [TestMethod]
        public void ValidDateTime2()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "TextToConvert", "5/1/2015 3:00 PM" }
            };

            //Expected value(s)
            DateTime expected1 = new DateTime(2015, 5, 1, 15, 0, 0);
            const bool expected2 = true;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test(s)
            Assert.AreEqual(expected1, output["ConvertedDate"]);
            Assert.AreEqual(expected2, output["IsValid"]);
        }

        [TestMethod]
        public void ValidDateTime3()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "TextToConvert", "5-1-2015 15:00:00" }
            };

            //Expected value(s)
            DateTime expected1 = new DateTime(2015, 5, 1, 15, 0, 0);
            const bool expected2 = true;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test(s)
            Assert.AreEqual(expected1, output["ConvertedDate"]);
            Assert.AreEqual(expected2, output["IsValid"]);
        }

        [TestMethod]
        public void ValidDateTime4()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "TextToConvert", "Fri, 08 May 2015 21:45:58" }
            };

            //Expected value(s)
            DateTime expected1 = new DateTime(2015, 5, 8, 21, 45, 58);
            const bool expected2 = true;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test(s)
            Assert.AreEqual(expected1, output["ConvertedDate"]);
            Assert.AreEqual(expected2, output["IsValid"]);
        }

        [TestMethod]
        public void ValidDateTime5()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "TextToConvert", "2009-06-15T13:45:30" }
            };

            //Expected value(s)
            DateTime expected1 = new DateTime(2009, 6, 15, 13, 45, 30);
            const bool expected2 = true;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test(s)
            Assert.AreEqual(expected1, output["ConvertedDate"]);
            Assert.AreEqual(expected2, output["IsValid"]);
        }

        [TestMethod]
        public void InValidDate()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "TextToConvert", "Hello World" }
            };

            //Expected value(s)
            DateTime expected1 = new DateTime(1, 1, 1, 0, 0, 0);
            const bool expected2 = false;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test(s)
            Assert.AreEqual(expected1, output["ConvertedDate"]);
            Assert.AreEqual(expected2, output["IsValid"]);
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
