using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using WindowsInput;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;
using System.Drawing;

namespace CommanderSpace
{
    public class CommandType
    {
        public string command;
        public string program;
        public string key;
        public string argument;
        public string performForAction; // double
        public string text; // double


        public List<int> figures;

        public CommandType()
        {
            argument = "";
            figures = new List<int>();
        }
    }

    public class Commander
    {
        public List<string> listPrograms = new List<string>();

        #region DllImport
        [DllImport("user32.dll")]
        static extern bool ShowWindow(int hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        public static extern int SetForegroundWindow(int hwnd);

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);
        #endregion

        public Commander() { }

        #region Check
        public bool IsOtherWord( string text )
        {
            if (text.Equals("to") ||
                text.Equals("the") ||
                text.Equals("of") ||
                text.Equals("and"))
            { return true; }
            return false;
        }
        public bool IsCommand( string text )
        {
            if (
                text.Equals("write") ||
                text.Equals("leftclick") ||
                text.Equals("rightclick") ||
                text.Equals("set") || 
                text.Equals("open") ||
                text.Equals("new") ||
                text.Equals("back") ||
                text.Equals("switchwindows") ||
                text.Equals("restore") ||
                text.Equals("default") ||
                text.Equals("normal") ||
                text.Equals("minimize") ||
                text.Equals("press") ||
                text.Equals("enter") ||
                text.Equals("maximize") ||
                text.Equals("hide") ||
                text.Equals("play") ||
                text.Equals("run") ||
                text.Equals("switch") ||
                text.Equals("next") ||
                text.Equals("go") ||
                text.Equals("show") ||
                text.Equals("close") ||
                text.Equals("run") ||
                text.Equals("hide") ||
                text.Equals("sleep") ||
                text.Equals("wait") ||
                text.Equals("select"))
            { return true; }
            return false;
        }
        public bool IsCommandWithoutArg( string text )
        {
            if (
                text.Equals("write") ||
                text.Equals("select") ||
                text.Equals("maximize") ||
                text.Equals("minimize") ||
                text.Equals("default") ||
                text.Equals("new") ||
                text.Equals("normal") ||
                text.Equals("next") ||
                text.Equals("close") ||
                text.Equals("back") ||
                text.Equals("switchwindows") ||
                text.Equals("switch") ||
                text.Equals("cancel"))
            { return true; }
            return false;
        }
        public bool IsProgram(string text)
        {
            if ( text.Equals("chrome") ||
                text.Equals("cmd") )
            { return true; }
            return false;
        }
        public bool IsKey( string text )
        {
            if ( text.Equals("space") ||
                 text.Equals("esc") ||
                 text.Equals("delete") )
            { return true; }
            return false;
        }
        public bool IsPerformForAction( string text )
        {
            if (text.Equals("double") )
            { return true; }
            return false;
        }
        #endregion
        #region WinFunctions
        public void Run( string programName, string programArgument )
        {
            try
            {
                if (programArgument == null)
                {
                    Process.Start(programName + ".exe");
                }
                else
                {
                    ProcessStartInfo procStartInfo = new ProcessStartInfo(programName + ".exe", programArgument);
                    Process p = new Process();
                    p.StartInfo = procStartInfo;
                    p.Start();
                }
            }
            catch { return; }
        }
        public void Show(string nameWindow)
        {
            if (nameWindow.Equals("window"))
            {
                ShowWindow(GetForegroundWindow().ToInt32(), 5);
            }
            else
            {
                foreach (var process in GetProcessWithName(nameWindow))
                {
                    ShowWindow(process.MainWindowHandle.ToInt32(), 5);
                    SetForegroundWindow(process.MainWindowHandle.ToInt32());
                }
            }
        }
        public void MaximizeWin(string nameWindow)
        {
            if (nameWindow.Equals("window") || nameWindow.Equals("size") || nameWindow.Equals("none"))
            {
                ShowWindow(GetForegroundWindow().ToInt32(), 3);
            }
            else
            {
                foreach (var process in GetProcessWithName(nameWindow))
                {
                    ShowWindow(process.MainWindowHandle.ToInt32(), 3);
                }
            }
        }
        public void MinimizeWin(string nameWindow)
        {
            if (nameWindow.Equals("window") || nameWindow.Equals("size") || nameWindow.Equals("none"))
            {
                ShowWindow(GetForegroundWindow().ToInt32(), 6);
            }
            else
            {
                foreach (var process in GetProcessWithName(nameWindow))
                {
                    ShowWindow(process.MainWindowHandle.ToInt32(), 6);
                }
            }
        }
        public void DefaultWin(string nameWindow)
        {
            int handle;
            if (nameWindow.Equals("window") || nameWindow.Equals("size") || nameWindow.Equals("none"))
            {
                Process p = GetForegroundProcess();
                if (p != null)
                {
                    handle = p.MainWindowHandle.ToInt32();
                    ShowWindow(handle, 10);
                }
            }
            else
            {
                foreach (Process p in GetProcessWithName(nameWindow))
                {
                    handle = p.MainWindowHandle.ToInt32();
                    ShowWindow(handle, 10);
                }
            }
        }

        public Process GetForegroundProcess()
        {
            int handleId = GetForegroundWindow().ToInt32();
            foreach (Process p in Process.GetProcesses())
            {
                if (p.MainWindowHandle.ToInt32().Equals(handleId))
                {
                    return p;
                }
                else if (p.MainWindowHandle.ToInt64().Equals(handleId))
                {

                }
            }
            return null;
        }
        public List<Process> GetProcessWithName(string name)
        {
            List<Process> listHandle = new List<Process>();
            foreach (Process p in Process.GetProcesses())
            {
                var procName = p.ProcessName;
                if (procName.Equals(name))
                {
                    listHandle.Add(p);
                }
            }
            return listHandle;
        }
        #endregion
        #region Mouse
        public void MouseLeftClick( Point cursor )
        {
            mouse_event((int)MouseEvent.MOUSEEVENTF_LEFTDOWN, cursor.X, cursor.Y, 0, 0);
            mouse_event((int)MouseEvent.MOUSEEVENTF_LEFTUP, cursor.X, cursor.Y, 0, 0);
        }
        public void MouseRightClick( Point cursor )
        {
            mouse_event((int)MouseEvent.MOUSEEVENTF_RIGHTDOWN, cursor.X, cursor.Y, 0, 0);
            mouse_event((int)MouseEvent.MOUSEEVENTF_RIGHTUP, cursor.X, cursor.Y, 0, 0);
        }
        public void Mouse( string mouseEvent, string performForAction, Point cursor )
        {
            switch (mouseEvent)
            {
                #region leftClick
                case "leftclick":
                    if (performForAction != null && performForAction.Equals("double"))
                    {
                        MouseLeftClick( cursor );
                        MouseLeftClick( cursor );
                    }
                    else
                    {
                        MouseLeftClick( cursor );
                    }
                break;
                #endregion
                #region RightClick
                case "rightclick":
                    if (performForAction != null && performForAction.Equals("double"))
                    {
                        MouseRightClick( cursor );
                        MouseRightClick( cursor );
                    }
                    else
                    {
                        MouseRightClick( cursor );
                    }
                break;
                #endregion
            }
        }
        #endregion

        public VirtualKeyCode StringToVirtualKey(string key)
        {
            switch (key)
            {
                case "space": return VirtualKeyCode.SPACE;
                case "delete": return VirtualKeyCode.DELETE;
            }
            return VirtualKeyCode.ACCEPT;
        }
        public CommandType ParseTextToCommand(string text)
        {
            string[] words = text.Split(' ');

            CommandType commandType = new CommandType();
            List<string> listWords = new List<string>(words);
            int val;


            for (int w = 0; w < listWords.Count; w++)
            {
                #region Other Word
                if (IsOtherWord(listWords[w]))
                {
                    listWords.RemoveAt(w);
                    w--;
                }
                #endregion
                #region Command
                else if (IsCommand(listWords[w]))
                {
                    commandType.command = listWords[w].ToLower();
                    if (IsCommandWithoutArg(listWords[w]))
                    {
                        commandType.argument = "none";
                    }
                }
                #endregion
                #region Program
                else if (IsProgram(listWords[w]))
                {
                    commandType.program = listWords[w];
                }
                #endregion
                #region Key
                else if (IsKey(listWords[w]))
                {
                    commandType.key = listWords[w];
                }
                #endregion
                #region Figure
                else if (Int32.TryParse(listWords[w], out val))
                {
                    commandType.figures.Add(val);
                }
                #endregion
                #region PerformForAction
                else if (IsPerformForAction(listWords[w]))
                {
                    commandType.performForAction = listWords[w].ToLower();
                }
                #endregion
                #region ELSE
                else if (commandType.argument != "none")
                {
                    commandType.argument += listWords[w] + " ";
                }
                #endregion
            }
            commandType.argument = commandType.argument.TrimEnd();
            commandType.text = text.Substring(6);
            return commandType;
        }
        public void Do(object obj)
        {
            CommandType commandType = (CommandType)obj;
            int x, y;
            Point cursorPos = new Point();

            switch (commandType.command)
            {
                #region open/run
                case "open":
                case "run":
                    if( commandType.program != null )
                    {
                        Run( commandType.program, commandType.argument );
                    }
                    else
                    {
                        Run( commandType.argument, null );
                    }
                    break;
                #endregion
                #region sleep/wait
                case "sleep":
                case "wait":
                    int val;
                    if ( Int32.TryParse( commandType.argument, out val ) )
                    {
                        Thread.Sleep( val );
                    }
                break;
                #endregion
                #region switch/next
                case "switch":
                    switch (commandType.argument)
                    {
                        case "window":
                            InputSimulator.SimulateKeyPress(VirtualKeyCode.TAB);
                            break;
                    }
                    break;
                case "next":
                    InputSimulator.SimulateKeyPress(VirtualKeyCode.TAB);
                break;
                #endregion
                #region go/show
                case "go":
                case "show":
                    if (commandType.program != null)
                    {
                        Show(commandType.program);
                    }
                    else if (commandType.argument != null)
                    {
                        Show(commandType.argument);
                    }
                    break;
                #endregion
                #region press
                case "press":
                    InputSimulator.SimulateKeyPress( StringToVirtualKey( commandType.key ) );
                break;
                #endregion
                #region write
                case "write":
                    InputSimulator.SimulateTextEntry( commandType.text );
                break;
                #endregion
                #region select
                case "select":
                    InputSimulator.SimulateKeyPress( VirtualKeyCode.RETURN );
                break;
                #endregion
                #region close
                case "close":
                    if ( commandType.program != null )
                    {
                        foreach ( var process in GetProcessWithName(commandType.program) )
                        { try { process.Kill(); } catch { return; } }
                    }
                    else if ( commandType.argument.Equals("window") || commandType.argument.Equals("none") )
                    {
                        try { GetForegroundProcess().Kill(); } catch { }
                    }
                    else
                    {
                        foreach (var process in GetProcessWithName(commandType.argument))
                        { try { process.Kill(); } catch { return; } }
                    }
                    break;
                #endregion
                #region hide/minimize
                case "minimize":
                case "hide":
                    if (commandType.program != null)
                    {
                        MinimizeWin(commandType.program);
                    }
                    else
                    {
                        MinimizeWin(commandType.argument);
                    }  
                    break;
                #endregion
                #region maximize
                case "maximize":
                    if (commandType.program != null)
                    {
                        MaximizeWin(commandType.program);
                    }
                    else
                    {
                        MaximizeWin(commandType.argument);
                    }
                    break;
                #endregion
                #region switchWindows
                case "switchwindows":
                    Process.Start("rundll32.exe", "DwmApi #105");
                    return;
                #endregion
                #region default/normal
                case "restore":
                case "default":
                case "normal":
                    if (commandType.program != null)
                    {
                        DefaultWin(commandType.program);
                    }
                    else
                    {
                        DefaultWin(commandType.argument);
                    }
                    return;
                #endregion
                #region set
                case "set":
                    switch( commandType.argument )
                    {
                        case "cursor":
                            #region inititalization
                            if ( commandType.figures.Count >= 2 )
                            {
                                x = commandType.figures[0];
                                y = commandType.figures[1];
                            }
                            else
                            {
                                x = Cursor.Position.X;
                                y = Cursor.Position.Y;
                            }
                            #endregion

                            Cursor.Position = new Point( x, y );
                        break;
                    }
                break;
                #endregion
                #region leftClick
                case "leftclick":
                    #region inititalization
                    if (commandType.figures.Count >= 2)
                    {
                        cursorPos.X = commandType.figures[0];
                        cursorPos.Y = commandType.figures[1];

                        Cursor.Position = new Point(cursorPos.X, cursorPos.Y);
                    }
                    else
                    {
                        x = Cursor.Position.X;
                        y = Cursor.Position.Y;
                    }
                    #endregion
                    
                    Mouse( commandType.command, commandType.performForAction, cursorPos );
                break;
                #endregion
                #region rightClick
                case "rightclick":
                    #region inititalization
                    if (commandType.figures.Count >= 2)
                    {
                        cursorPos.X = commandType.figures[0];
                        cursorPos.Y = commandType.figures[1];

                        Cursor.Position = new Point(cursorPos.X, cursorPos.Y);
                    }
                    else
                    {
                        x = Cursor.Position.X;
                        y = Cursor.Position.Y;
                    }
                    #endregion

                    Mouse(commandType.command, commandType.performForAction, cursorPos);
                break;
                #endregion
            }
        }
    }
}


enum MouseEvent
{
    MOUSEEVENTF_LEFTDOWN = 0x02,
    MOUSEEVENTF_LEFTUP = 0x04,
    MOUSEEVENTF_RIGHTDOWN = 0x08,
    MOUSEEVENTF_RIGHTUP = 0x10,
}