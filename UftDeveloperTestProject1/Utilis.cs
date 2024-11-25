using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UftDeveloperTestProject1
{
    public static class Utilis
    {
        public static void LogToFile(string message)
        {
            string logFilePath = EnvironmentSetting.pathToLogFile; // Path to your log file

            using (StreamWriter writer = new StreamWriter(logFilePath, true))
            {
                writer.WriteLine($"{DateTime.Now}: {message}");
            }
        }

        public static void launchExe(string exePath)
        {


            // Create a new process
            Process appProcess = new Process { StartInfo = { FileName = exePath } };
            appProcess.Start();
        }
    }
}
