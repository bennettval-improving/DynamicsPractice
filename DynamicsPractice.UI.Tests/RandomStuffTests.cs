using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Configuration;
using System.Security;

namespace DynamicsPractice.UI.Tests
{
    [TestClass]
    public class RandomStuffTests
    {
        private SecureString _username = ConfigurationManager.AppSettings["TestEmail"].ToSecureString();
        private SecureString _password = ConfigurationManager.AppSettings["TestPassword"].ToSecureString();
        private Uri _tenantUrl = new Uri(ConfigurationManager.AppSettings["TenantUrl"]);

        [TestMethod]
        public void SuccessfulSave_WhenAllFieldsFilledOut()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_tenantUrl, _username, _password);

                xrmApp.Navigation.OpenApp(AppName.Sales);

                xrmApp.Navigation.OpenSubArea("Sales", "Random Stuff");

                xrmApp.CommandBar.ClickCommand("New");

                xrmApp.Entity.SetValue("uitest_name", TestSettings.GetRandomString(5, 15));

                xrmApp.Entity.SelectTab("Super Important");

                xrmApp.Entity.SetValue(GetLookupItem("uitest_contact", "Nancy Anderson (sample)"));

                xrmApp.Entity.SetValue("uitest_superpowerdescription", "Can Sleep for long periods of time.");

                xrmApp.Entity.Save();
            }
        }

        [TestMethod]
        public void UnSuccessfulSave_WhenOnlyNameIsFilledOut()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_tenantUrl, _username, _password);
                xrmApp.Navigation.OpenApp(AppName.Sales);

                xrmApp.Navigation.OpenSubArea("Sales", "Random Stuff");
                var expectedCount = xrmApp.Grid.GetGridItems().Count;
                xrmApp.CommandBar.ClickCommand("New");

                xrmApp.Entity.SetValue("uitest_name", TestSettings.GetRandomString(5, 15));
                xrmApp.Entity.Save();

                var notifications = xrmApp.Entity.GetFormNotifications();

                xrmApp.Navigation.OpenSubArea("Sales", "Random Stuff");
                xrmApp.Dialogs.ConfirmationDialog(false);
                var acutalCount = xrmApp.Grid.GetGridItems().Count;

                Assert.IsTrue(notifications.Count == 2);
                Assert.AreEqual(expectedCount, acutalCount);
            }
        }

        private LookupItem GetLookupItem(string name, string value)
        {
            return new LookupItem { Name = name, Value = value };
        }
    }
}
