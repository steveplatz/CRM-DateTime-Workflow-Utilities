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
    public class ToStringTests
    {
        #region Class Constructor
        private readonly string _namespaceClassAssembly;
        public ToStringTests()
        {
            //[Namespace.class name, assembly name] for the class/assembly being tested
            //Namespace and class name can be found on the class file being tested
            //Assembly name can be found under the project properties on the Application tab
            _namespaceClassAssembly = "LAT.WorkflowUtilities.DateTimes.ToString" + ", " + "LAT.WorkflowUtilities.DateTimes";
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
        public void Date_Default()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "DateToUse", new DateTime(2014, 12, 31, 23, 0, 0, 0)},
                { "EvaluateAsUserLocal", true},
                { "Culture", "en-US" },
                { "Format", "g" }
            };

            //Expected value
            const string expected = "12/31/2014 11:00 PM";

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, ValidDayUserLocalSetup);

            //Test
            Assert.AreEqual(expected, output["FormattedDateString"]);
        }

        [TestMethod]
        public void Date_esES()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "DateToUse", new DateTime(2014, 12, 31, 23, 0, 0, 0)},
                { "EvaluateAsUserLocal", true},
                { "Culture", "es-ES" },
                { "Format", "g" }
            };

            //Expected value
            const string expected = "31/12/2014 23:00";

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, ValidDayUserLocalSetup);

            //Test
            Assert.AreEqual(expected, output["FormattedDateString"]);
        }

        [TestMethod]
        public void Date_InvalidCulture()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "DateToUse", new DateTime(2014, 12, 31, 23, 0, 0, 0)},
                { "EvaluateAsUserLocal", true},
                { "Culture", "xx5x" },
                { "Format", "g" }
            };

            //Expected value
            const string expected = null;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, ValidDayUserLocalSetup);

            //Test
            Assert.AreEqual(expected, output["FormattedDateString"]);
        }

        [TestMethod]
        public void Date_InvalidFormat()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "DateToUse", new DateTime(2014, 12, 31, 23, 0, 0, 0)},
                { "EvaluateAsUserLocal", true},
                { "Culture", "en-US" },
                { "Format", "ZZZ" }
            };

            //Expected value
            const string expected = "ZZZ";

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, ValidDayUserLocalSetup);

            //Test
            Assert.AreEqual(expected, output["FormattedDateString"]);
        }

        [TestMethod]
        public void Date_ValidDateUtc()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "DateToUse", new DateTime(2015, 1, 1, 5, 0, 0, 0, DateTimeKind.Utc)},
                { "EvaluateAsUserLocal", false},
                { "Culture", "en-US" },
                { "Format", "g" }
            };

            //Expected value
            const string expected = "1/1/2015 5:00 AM";

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, ValidDayUserLocalSetup);

            //Test
            Assert.AreEqual(expected, output["FormattedDateString"]);
        }

        [TestMethod]
        public void Date_ValidDateUserLocal()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "DateToUse", new DateTime(2015, 1, 1, 5, 0, 0, 0, DateTimeKind.Utc)},
                { "EvaluateAsUserLocal", true},
                { "Culture", "en-US" },
                { "Format", "g" }
            };

            //Expected value
            const string expected = "12/31/2014 11:00 PM";

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, ValidDayUserLocalSetup);

            //Test
            Assert.AreEqual(expected, output["FormattedDateString"]);
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

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> ValidDayUserLocalSetup(Mock<IOrganizationService> serviceMock)
        {
            EntityCollection userSettings = new EntityCollection();
            Entity userSetting = new Entity("usersettings");
            userSetting["timezonecode"] = 20;
            userSettings.Entities.Add(userSetting);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(userSettings);

            OrganizationResponse localTime = new OrganizationResponse();
            ParameterCollection results = new ParameterCollection { { "LocalTime", new DateTime(2014, 12, 31, 23, 0, 0) } };
            localTime.Results = results;

            serviceMock.Setup(t =>
                t.Execute(It.IsAny<OrganizationRequest>()))
                .ReturnsInOrder(localTime);

            return serviceMock;
        }
    }
}
