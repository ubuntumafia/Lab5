using System;
using System.IO;

public static class logger
{
    private static StreamWriter logWriter;

    public static void InitializeLogger(string logFileName, bool append)
    {
        if (append)
        {
            logWriter = new StreamWriter(new FileStream(logFileName, FileMode.Append, FileAccess.Write))
            {
                AutoFlush = true
            };
        }
        else
        {
            string newLogFileName = GetNewLogFileName(logFileName);
            logWriter = new StreamWriter(new FileStream(newLogFileName, FileMode.Create, FileAccess.Write))
            {
                AutoFlush = true
            };
        }
    }

    private static string GetNewLogFileName(string baseFileName)
    {
        int fileIndex = 1;
        string newLogFileName;

        do
        {
            newLogFileName = $"{baseFileName}{fileIndex}.txt";
            fileIndex++;
        } while (File.Exists(newLogFileName));

        return newLogFileName;
    }

    public static void LogAction(string message)
    {
        var logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
        logWriter.WriteLine(logMessage);
    }

    public static void CloseLogger()
    {
        logWriter?.Close();
    }
}
