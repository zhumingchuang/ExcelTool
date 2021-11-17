using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelTool
{
    public enum Level
    {
        ERROR,
        INFO,
        DEBUG,
        ALL
    }
    public class Log
    {
        private static Level LOG_LEVENL;

        private static bool IsTrackWrite;

        private static bool IsDisableSingleFileMode;

        private const long maxStreamSize = 3145728L;

        private readonly static object SyncRoot;

        private static string _LogFileCreateDate;

        private static string LogFileName;

        private static Stream _safeStream;

        private static FileStream _fileStream;

        private static string _LogFileFullName;

        private static string LogFileSaveDirPath;

        private static Stream FileStream
        {
            get
            {
                return Log._safeStream;
            }
        }

        public static string LogFileFullPath
        {
            get
            {
                if (string.IsNullOrEmpty(Log._LogFileFullName))
                {
                    Log._LogFileFullName = string.Format("{0}/Logs/{1}", Log.LogFileSaveDirPath.TrimEnd(new char[] { '/' }), Log.LogFileName);
                }
                return Log._LogFileFullName;
            }
        }

        private static string NowTimeDate
        {
            get
            {
                return DateTime.Now.ToString("yyyyMMdd");
            }
        }

        static Log()
        {
            Log.LOG_LEVENL = Level.ALL;
            Log.IsTrackWrite = true;
            Log.IsDisableSingleFileMode = false;
            Log.SyncRoot = new object();
            Log._LogFileCreateDate = string.Empty;
            Log.LogFileName = "Debug.log";
            Log.LogFileSaveDirPath = Application.StartupPath;
        }

        public Log()
        {
        }

        public static void ApplySetting(Level level, string logFileSaveDirPath, bool bTrackWrite, bool isDisableSingleFileMode)
        {
            Log.LOG_LEVENL = level;
            Log.LogFileSaveDirPath = logFileSaveDirPath;
            Log.IsTrackWrite = bTrackWrite;
            Log.IsDisableSingleFileMode = isDisableSingleFileMode;
        }

        private static void BufferWriteToFile(byte[] buffer, int size)
        {
            bool flag;
            try
            {
                lock (Log.SyncRoot)
                {
                    if (!Log.IsDisableSingleFileMode)
                    {
                        flag = false;
                    }
                    else
                    {
                        flag = (Log.FileStream.Length > (long)3145728 ? true : Log._LogFileCreateDate != Log.NowTimeDate);
                    }
                    if (flag)
                    {
                        Log.CloseFile();
                        Log.SaveAs();
                        Log.CreateOrOpenFile();
                    }
                    Log.FileStream.Write(buffer, 0, size);
                    Log.FileStream.Flush();
                }
            }
            catch (Exception exception)
            {
                Log.TrackLog(exception.ToString());
            }
        }

        private static void CloseFile()
        {
            if (Log._fileStream != null)
            {
                Log._safeStream.Flush();
                Log._safeStream.Close();
                Log._safeStream.Dispose();
                Log._fileStream.Dispose();
                Log._safeStream = null;
                Log._fileStream = null;
            }
        }

        private static void CreateDirectory()
        {
            FileInfo fileInfo = new FileInfo(Log.LogFileFullPath);
            if (!fileInfo.Directory.Exists)
            {
                fileInfo.Directory.Create();
            }
        }

        private static string CreateNewFilePath()
        {
            int num = 1;
            int num1 = Log.LogFileName.LastIndexOf('.');
            string str = string.Format("{0}/Logs/{1}", Log.LogFileSaveDirPath, Log.LogFileName);
            DateTime creationTime = (new FileInfo(str)).CreationTime;
            while (File.Exists(str))
            {
                str = string.Format("{0}/Logs/{1}", Log.LogFileSaveDirPath, Log.LogFileName.Insert(num1, string.Format("_{0:yyyy_MM_dd}_{1}", creationTime, num.ToString("D2"))));
                num++;
            }
            return str;
        }

        private static void CreateOrOpenFile()
        {
            Log.CreateDirectory();
            Log._fileStream = new FileStream(Log.LogFileFullPath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
            Log._safeStream = Stream.Synchronized(Log._fileStream);
            Log._LogFileCreateDate = Log.NowTimeDate;
        }

        private static string GetFormatStr(string Str, bool PreCRLF)
        {
            StackFrame frame = (new StackTrace()).GetFrame(3);
            return string.Format("{0}[{1:HH:mm:ss.fff}] [{2:x8}] [{3}.{4}] | {5}\r\n", new object[] { (PreCRLF ? "\r\n" : ""), DateTime.Now, Thread.CurrentThread.ManagedThreadId, frame.GetMethod().ReflectedType.Name, frame.GetMethod().Name, Str });
        }

        private static void SaveAs()
        {
            File.Copy(Log.LogFileFullPath, Log.CreateNewFilePath(), true);
        }

        public static void SetFileName(string fileName)
        {
            Log.LogFileName = fileName;
        }

        public static void SetLogFileSaveDirPath(string dirPath)
        {
            Log.LogFileSaveDirPath = dirPath;
        }

        public static void SetLogLevel(Level level)
        {
            Log.LOG_LEVENL = level;
        }

        public static void Shutdown()
        {
            try
            {
                Log.CloseFile();
            }
            catch (Exception exception)
            {
                Log.TrackLog(exception.ToString());
            }
        }

        private static void TrackLog(string str)
        {
            if (Log.IsTrackWrite)
            {
                Debug.WriteLine(str);
                Trace.Write(str);
            }
        }

        public static void Write(Level level, string content, bool preCRLF = false)
        {
            try
            {
                if (level <= Log.LOG_LEVENL)
                {
                    Log.WriteFile(content, preCRLF);
                }
            }
            catch
            {
            }
        }

        public static void Write(Level level, object content, bool preCRLF = false)
        {
            try
            {
                if (level <= Log.LOG_LEVENL)
                {
                    Log.WriteFile(content.ToString(), preCRLF);
                }
            }
            catch
            {
            }
        }

        private static void WriteFile(string content, bool preCRLF)
        {
            if (Log.FileStream == null)
            {
                if ((!Log.IsDisableSingleFileMode ? false : File.Exists(Log.LogFileFullPath)))
                {
                    Log.SaveAs();
                }
                Log.CreateOrOpenFile();
                Log.TrackLog(Log.GetFormatStr(string.Concat("Log Services Start, FilePath = ", Log.LogFileFullPath), false));
            }
            string formatStr = Log.GetFormatStr(content, preCRLF);
            Log.TrackLog(formatStr);
            byte[] bytes = (new UTF8Encoding()).GetBytes(formatStr);
            Log.BufferWriteToFile(bytes, (int)bytes.Length);
        }
    }
}
