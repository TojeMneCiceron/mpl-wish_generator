using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Data data = new Data();

            var app_doc = new Microsoft.Office.Interop.Word.Application();
            var config = @"C:\Users\Пользователь\Desktop\config.xlsx";
            data.ReadData(config);

            WishGenerator wishGen = new WishGenerator();
            wishGen.Wishes = data.Wishes;
            wishGen.Generate(data.Names.Count);

            Postcard pCard = new Postcard();
            pCard.CreatePostcard(data.Template, data.Font, data.Names, wishGen.Generated);

            Console.WriteLine("Генерация завершена!");

            data.Close();
            
            Console.ReadKey();
        }
    }
}