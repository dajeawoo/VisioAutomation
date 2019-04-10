using System.Drawing;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VisioAutomation_Tests.Scripting
{
    [TestClass]
    public class ScriptingApplicationTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Test_Resize_Application_Window1()
        {
            var active_app = new VisioScripting.TargetActiveApplication();

            var desired_size = new System.Drawing.Size(600, 700);
            var client = this.GetScriptingClient();
            var old_rect = client.Application.GetWindowRectangle(active_app);
            var new_rect = new System.Drawing.Rectangle(old_rect.X, old_rect.Y, desired_size.Width, desired_size.Height);

            client.Application.SetWindowRectangle(active_app, new_rect);
            var actual_rect1 = client.Application.GetWindowRectangle(active_app);
            Assert.AreEqual(desired_size, actual_rect1.Size);

            client.Application.SetWindowRectangle(active_app, old_rect);
            var actual_rect2 = client.Application.GetWindowRectangle(active_app);
            Assert.AreEqual(old_rect.Size, actual_rect2.Size);
            Assert.AreEqual(old_rect, actual_rect2);

        }

        [TestMethod]
        public void Scripting_Test_Resize_Application_Window2()
        {
            var client = this.GetScriptingClient();
            var page_size = new VisioAutomation.Geometry.Size(10,5);
            var doc = client.Document.NewDocument(page_size);

            var page = client.Page.GetActivePage();
            var tagetpage = new VisioScripting.TargetPage(page);

            var pagesize = client.Page.GetPageSize(tagetpage);
            Assert.AreEqual(10.0, pagesize.Width);
            Assert.AreEqual(5.0, pagesize.Height);

            var targetwindow = new VisioScripting.TargetWindow();

            Assert.AreEqual(0, client.Selection.GetSelection(targetwindow).Count);
            client.Draw.DrawRectangle(1, 1, 2, 2);
            Assert.AreEqual(1, client.Selection.GetSelection(targetwindow).Count);

            var targetdoc = new VisioScripting.TargetDocument();
            client.Document.CloseDocument(targetdoc, true);
        }

        [TestMethod]
        public void Scripting_Test_App_to_Front()
        {
            var client = this.GetScriptingClient();
            var activeapp = new VisioScripting.TargetActiveApplication();
            client.Application.MoveWindowToFront(activeapp);
        }

        [TestMethod]
        public void Scripting_Undo_Scenarios()
        {
            var client = this.GetScriptingClient();
            var page_size = new VisioAutomation.Geometry.Size(8.5,11);
            var drawing = client.Document.NewDocument(page_size);

            var targetdoc = new VisioScripting.TargetDocument();
            var page = client.Page.NewPage(targetdoc, page_size, false);
            Assert.AreEqual(0, page.Shapes.Count);
            page.DrawRectangle(1, 1, 3, 3);
            Assert.AreEqual(1, page.Shapes.Count);
            var activeapp = new VisioScripting.TargetActiveApplication();
            client.Undo.UndoLastAction(activeapp);
            Assert.AreEqual(0, page.Shapes.Count);
            client.Document.CloseDocument(targetdoc, true);
        }

        [TestMethod]
        public void Scripting_CloseDocument_Scenarios()
        {
            var page_size = new VisioAutomation.Geometry.Size(8.5, 11);
            var client = this.GetScriptingClient();
            var doc1 = client.Document.NewDocument(page_size);
            var doc2 = client.Document.NewDocument(page_size);
            var doc3 = client.Document.NewDocument(page_size);

            client.Document.CloseAllDocumentsWithoutSaving();

            Assert.IsFalse(client.Document.HasActiveDocument);
            var application = client.Application.GetAttachedApplication();
            var documents = application.Documents;
            Assert.AreEqual(0, documents.Count);
        }
    }
}