using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Times
{
    class Times
    {
        static void Main(string[] args)
        {
            using (FileStream fs = File.Create("times.docx")) 
            {
                var random = new Random();

                WordDocument doc = new WordDocument();
                doc.EnsureMinimal();

                WParagraphStyle style = doc.Styles.FindByName("Normal") as WParagraphStyle;
                style.CharacterFormat.FontName = "Ariel";
                style.CharacterFormat.FontSize = 14;


                doc.LastParagraph.AppendText("Hi, hi, good morning");

                var section = doc.LastSection;

                section.AddParagraph();

                for (int i = 0; i < 20; i++)
                {
                    int x = random.Next(1, 12);
                    int y = random.Next(1, 12);

                    var paragraph = section.AddParagraph();
                    paragraph.AppendText(String.Format("\t{0} x {1} =", x, y));
                    section.AddParagraph();
                }

                doc.Save(fs, FormatType.Docx);
                doc.Close();
            }
        }
    }
}