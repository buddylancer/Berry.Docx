using Microsoft.VisualStudio.TestTools.UnitTesting;
using Berry.Docx;
using Berry.Docx.Documents;

namespace Berry.Docx.Tests
{
    [TestClass]
    public class ParagrphTest
    {
        [TestMethod]
        [DataRow("���Ƕ���1")]
        [DataRow("���Ƕ���2")]
        [DataRow("���Ƕ���3")]
        public void Set_ParagraphText_ReturnSameText(string text)
        {
            var doc = new Document("test.docx");
            Paragraph p = new Paragraph(doc) { Text = text };
            Assert.IsTrue(text == p.Text);
            doc.Close();
        }
    }
}
