using System;
using System.IO;

namespace SeleniumAutomation
{
    public static class WriteFile
    {
        
        public static void WriteToFile(string outputtext)
        {
            FileStream writepath;
            StreamWriter fileWriter;
            TextWriter oldOut = Console.Out;

            outputtext = DateTime.Now.ToString() + " - " + outputtext;

            try
            {
                FileMode mode = FileMode.Append;
                if (!File.Exists("C:\\Selenium Result\\Test Results.txt"))
                {
                    mode = FileMode.OpenOrCreate;
                }
                writepath = new FileStream("C:\\Selenium Result\\Test Results.txt", mode, FileAccess.Write);
                fileWriter = new StreamWriter(writepath);
                fileWriter.WriteLine(outputtext);
            }
            catch(Exception e)
            {
                Console.WriteLine("Cannont open the writepath file");
                Console.WriteLine(e.Message);
                return;
            }
           
                Console.SetOut(fileWriter);
                Console.WriteLine("This test has passed");
                Console.SetOut(oldOut);

                fileWriter.Close();
                writepath.Close();
            Console.WriteLine("DONE");
            

        }

    }
}
