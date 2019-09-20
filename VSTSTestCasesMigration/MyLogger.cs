using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VSTSTestCasesMigration
{
    public class MyLogger
    {
        public static string FileName { get; set; }

        public static void Log(string message)
        {
            FileStream fileStream;
            StreamWriter writer;
            TextWriter oldOut = Console.Out;
            try
            {
                fileStream = new FileStream(FileName, FileMode.Append, FileAccess.Write);
                writer = new StreamWriter(fileStream);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Cannot open " + FileName + " for writing");
                Console.WriteLine(ex.Message);
                return;
            }

            Console.SetOut(writer);
            Console.WriteLine(message);
            Console.SetOut(oldOut);
            writer.Close();
            fileStream.Close();
        }
    }
}
