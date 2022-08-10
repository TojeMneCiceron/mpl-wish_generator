using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ConsoleApp1
{
    class Data
    {
        Application app;
        Workbook book;
        List<string> names;
        List<List<string>> wishes;
        string template;
        string font;
        public Data()
        {
            app = new Microsoft.Office.Interop.Excel.Application();
            book = null;
            names = new List<string>();
            wishes = new List<List<string>>();
        }
        public List<string> Names
        {
            get { return names; }
        }
        public List<List<string>> Wishes
        {
            get { return wishes; }
        }
        public string Template
        {
            get { return template; }
        }
        public string Font
        {
            get { return font; }
        }
        public void Close()
        {
            book.Close();
            app.Quit();
            Console.ReadKey();
            System.Environment.Exit(0);
        }
        int AllWishCount()
        {
            int count = 0;
            int n = wishes.Count;

            for (int i = 0; i < n - 2; i++)
                for (int j = i + 1; j < n - 1; j++)
                    for (int k = j + 1; k < n; k++)
                        count += wishes[i].Count * wishes[j].Count * wishes[k].Count;

            return count;
        }
        public void ReadData(string path)
        {
            try
            {
                book = app.Workbooks.Add(path);
            }
            catch (Exception e)
            {
                Console.WriteLine(path);
                Console.WriteLine(e.Message);
                Close();
                Console.ReadKey();
                return;
            }
            Worksheet sheet;
            sheet = app.Worksheets["config"];           //config
            if (sheet.Cells[2, 1].Value2 != null)
                template = sheet.Cells[2, 1].Value2;
            else
            {
                Console.WriteLine("Укажите имя шаблона");
                Close();
            }
            if (sheet.Cells[2, 2].Value2 != null)
                font = sheet.Cells[2, 2].Value2;
            else
                font = "Arial";

            sheet = app.Worksheets["names"];            //names
            for (int i = 1; sheet.Cells[i, 1].Value2 != null; i++)
                names.Add(sheet.Cells[i, 1].Value2);
            if (names.Count == 0)
            {
                Console.WriteLine("Добавьте не менее 1 имени");
                Close();
            }

            sheet = app.Worksheets["wishes"];           //wishes
            for (int i = 1; sheet.Cells[1, i].Value2 != null; i++)
            {
                List<string> wish_group = new List<string>();
                for (int j = 2; sheet.Cells[j, i].Value2 != null; j++)
                    wish_group.Add(sheet.Cells[j, i].Value2);
                wishes.Add(wish_group);
            }
            if (wishes.Count < 3)
            {
                Console.WriteLine("Добавьте не менее трех групп пожеланий");
                Close();
            }
            if (names.Count > AllWishCount())
            {
                Console.WriteLine("Пополните список пожеланий");
                Close();
            }
        }
    }
}
