using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
    class Program
    {
        /*static void Main(string[] args)
        {

            //string sLine = String.Empty;

            //FileStream fs = new FileStream(@"C:\test\123.docx", FileMode.Open, FileAccess.Read);

            ////StreamReader sr = new StreamReader(fs, System.Text.UTF8Encoding.Default);
            //StreamReader sr = new StreamReader(fs);

            //sLine = sr.ReadToEnd();

            //Document doc = new Document();
            //doc.LoadFromFile(@"C:\test\1234.docx");


            Application application = new Application();
            Document document = application.Documents.Open(@"C:\test\1234.docx");
            string allText = document.Content.Text;
            application.Quit();

            // Loop through all words in the document.
            // count = document.Words.Count;
            ////string str = document.;
            //for (int i = 1; i <= count; i++)
            //{
            //    // Write the word.
            //    string text = document.Words[i].Text;
            //    Console.WriteLine("Word {0} = {1}", i, text);
            //}
            //// Close word.

        }*/


         private void readFileContent(string path)
        {
            //TextExtractor extractor = new TextExtractor(path);
            //string text = extractor.ExtractText();
            //Console.WriteLine(text);
        }
        static void Main(string[] args)
        {
            Program cs = new Program();
            string path = @"C:\test\1234.docx";
            cs.readFileContent(path);
            Console.ReadLine();
        }
    }
}

