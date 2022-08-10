using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using System.IO;

namespace ConsoleApp1
{
    class Postcard
    {
        Microsoft.Office.Interop.Word.Application app;
        Document doc;
        public Postcard()
        {
            app = new Microsoft.Office.Interop.Word.Application();
            doc = null;
        }
        public void Close()
        {
            doc.Close();
            app.Quit();
        }
        void Save()
        {
            string dir = Directory.GetCurrentDirectory() + "\\saved postcards";
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            int fileNum = 1;
            while (File.Exists(dir + "\\postcard" + fileNum.ToString() + ".docx"))
                fileNum++;
            doc.SaveAs2(dir + "\\postcard" + fileNum.ToString() + ".docx");
        }
        public void CreatePostcard(string template, string font, List<string> names, List<List<string>> wishes)
        {
            try
            {
                doc = app.Documents.Add(template);
            }
            catch (Exception e)
            {
                Console.WriteLine(template);
                Console.WriteLine(e.Message);
                app.Quit();
                Console.ReadKey();
                return;
            }

            for (int i = 0; i < names.Count; i++)
            {

                doc.Bookmarks["Name"].Range.Text = names[i];

                doc.Bookmarks["Wish1"].Range.Text = wishes[i][0];
                doc.Bookmarks["Wish2"].Range.Text = wishes[i][1];
                doc.Bookmarks["Wish3"].Range.Text = wishes[i][2];

                if (i != names.Count - 1)
                {
                    app.Selection.EndKey(WdUnits.wdStory);
                    app.Selection.InsertNewPage();
                    app.Selection.InsertFile(template, "", true, false, false);
                    app.Selection.EndKey(WdUnits.wdStory);
                    //app.Selection.Delete(1);
                }
            }

            app.Selection.WholeStory();
            app.Selection.Font.Name = font;

            //app.Visible = true;

            Save();
            Close();
        }
    }
}
