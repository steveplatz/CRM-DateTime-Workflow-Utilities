﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
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
    public class GetQuarterStartEndTests
    {
        #region Class Constructor
        private readonly string _namespaceClassAssembly;
        public GetQuarterStartEndTests()
        {
            //[Namespace.class name, assembly name] for the class/assembly being tested
            //Namespace and class name can be found on the class file being tested
            //Assembly name can be found under the project properties on the Application tab
            _namespaceClassAssembly = "LAT.WorkflowUtilities.DateTimes.GetQuarterStartEnd" + ", " + "LAT.WorkflowUtilities.DateTimes";
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
        public void ValidDate1UserLocal()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "DateToUse", new DateTime(2015, 1, 1, 5, 0, 0, DateTimeKind.Utc) },
                { "EvaluateAsUserLocal", true } 
            };

            //Expected value(s)
            DateTime expected1 = new DateTime(2014, 10, 1, 0, 0, 0);
            DateTime expected2 = new DateTime(2014, 12, 31, 23, 59, 59, 999);

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, ValidDate1UserLocalSetup);

            //Test(s)
            Assert.AreEqual(expected1, output["QuarterStartDate"]);
            Assert.AreEqual(expected2, output["QuarterEndDate"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> ValidDate1UserLocalSetup(Mock<IOrganizationService> serviceMock)
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

        [TestMethod]
        public void ValidDate2UserLocal()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "DateToUse", new DateTime(2015, 2, 1, 10, 0, 0, DateTimeKind.Utc) },
                { "EvaluateAsUserLocal", true } 
            };

            //Expected value(s)
            DateTime expected1 = new DateTime(2015, 1, 1, 0, 0, 0);
            DateTime expected2 = new DateTime(2015, 3, 31, 23, 59, 59, 999);

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, ValidDate2UserLocalSetup);

            //Test(s)
            Assert.AreEqual(expected1, output["QuarterStartDate"]);
            Assert.AreEqual(expected2, output["QuarterEndDate"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> ValidDate2UserLocalSetup(Mock<IOrganizationService> serviceMock)
        {
            EntityCollection userSettings = new EntityCollection();
            Entity userSetting = new Entity("usersettings");
            userSetting["timezonecode"] = 20;
            userSettings.Entities.Add(userSetting);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(userSettings);

            OrganizationResponse localTime = new OrganizationResponse();
            ParameterCollection results = new ParameterCollection { { "LocalTime", new DateTime(2015, 2, 1, 4, 0, 0) } };
            localTime.Results = results;

            serviceMock.Setup(t =>
                t.Execute(It.IsAny<OrganizationRequest>()))
                .ReturnsInOrder(localTime);

            return serviceMock;
        }

        [TestMethod]
        public void ValidDate2Utc()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object>
            {
                { "DateToUse", new DateTime(2015, 1, 1, 5, 0, 0, DateTimeKind.Utc) },
                { "EvaluateAsUserLocal", false } 
            };

            //Expected value(s)
            DateTime expected1 = new DateTime(2015, 1, 1, 0, 0, 0);
            DateTime expected2 = new DateTime(2015, 3, 31, 23, 59, 59, 999);

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, null);

            //Test(s)
            Assert.AreEqual(expected1, output["QuarterStartDate"]);
            Assert.AreEqual(expected2, output["QuarterEndDate"]);
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
