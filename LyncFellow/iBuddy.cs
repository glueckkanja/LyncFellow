using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;

namespace LyncFellow
{
    class iBuddy
    {
        [StructLayout(LayoutKind.Sequential)]
        public struct SECURITY_ATTRIBUTES
        {
            public int nLength;
            public int lpSecurityDescriptor;
            public int bInheritHandle;
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern int CreateFile(string lpFileName, uint dwDesiredAccess, uint dwShareMode, ref SECURITY_ATTRIBUTES lpSecurityAttributes, int dwCreationDisposition, uint dwFlagsAndAttributes, int hTemplateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern int WriteFile(int hFile, ref byte lpBuffer, int nNumberOfBytesToWrite, ref int lpNumberOfBytesWritten, int lpOverlapped);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern int CloseHandle(int hObject);

        const int _threadWaitTime = 50;

        public enum Turn { Off = 3, Left = 1, Right = 2 }
        public enum Flap { Off = 3, Up = 1, Down = 2 }
        public enum Color { Off = 7, Red = 6, Green = 5, Blue = 3, Yellow = 4, Lila = 2, Cyan = 1, LightBlue = 0 }

        public bool IsAlive = false;
        public int LastWin32Error = 0;
        private int _hidFileHandle = -1;
        private Thread _workerThread = null;
        private bool _shutdown = false;

        private Color _currentColor = Color.Off;
        private Color _targetColor = Color.Off;
        private bool _heart = false;
        private Turn _turn = Turn.Off;
        private Flap _flap = Flap.Off;

        private uint _danceTimeout = 0;
        private uint _rainbowTimeout = 0;
        private uint _heartbeatTimeout = 0;
        private uint _flapTimeout = 0;

        public bool Init(string HidDevicePath)
        {
            SECURITY_ATTRIBUTES structure = new SECURITY_ATTRIBUTES();
            structure.lpSecurityDescriptor = 0;
            structure.bInheritHandle = Convert.ToInt32(true);
            structure.nLength = Marshal.SizeOf(structure);

            _hidFileHandle = CreateFile(HidDevicePath, 0xc0000000, 3, ref structure, 3, 0, 0);
            if (_hidFileHandle == -1)
            {
                LastWin32Error = Marshal.GetLastWin32Error();
                Trace.WriteLine(string.Format("LyncFellow: CreateFile result=-1, GetLastWin32Error()={0}", LastWin32Error));
                return false;
            }

            _workerThread = new Thread(Worker);
            _workerThread.Start();

            IsAlive = true;
            return true;
        }

        public void Release()
        {
            if (_workerThread != null)
            {
                _shutdown = true;
                _workerThread.Join();
                _workerThread = null;
            }

            if (IsAlive)
            {
                _currentColor = Color.Off;
                _heart = false;
                _turn = Turn.Off;
                _flap = Flap.Off;
                UpdateState();
                IsAlive = false;
            }

            if (_hidFileHandle != -1)
            {
                CloseHandle(_hidFileHandle);
                _hidFileHandle = -1;
            }

        }

        ~iBuddy()
        {
            Release();
        }

        private void Worker()
        {
            //int loopCounter = 0;

            while (IsAlive && !_shutdown)
            {
                bool updateState = false;

                if (_danceTimeout > 0)
                {
                    _danceTimeout--;
                    if (_danceTimeout == 0)
                    {
                        _turn = Turn.Off;
                    }
                    else if ((_danceTimeout % (300 / _threadWaitTime)) == 0)
                    {
                        _turn = _turn == Turn.Left ? Turn.Right : Turn.Left;
                    }
                    updateState = true;
                }

                if (_flapTimeout > 0)
                {
                    _flapTimeout--;
                    if (_flapTimeout == 0)
                    {
                        _flap = Flap.Off;
                    }
                    else if ((_flapTimeout % (300 / _threadWaitTime)) == 0)
                    {
                        _flap = _flap == Flap.Up ? Flap.Down : Flap.Up;
                    }
                    updateState = true;
                }

                if (_heartbeatTimeout > 0)
                {
                    _heartbeatTimeout--;
                    if (_heartbeatTimeout == 0)
                    {
                        _heart = false;
                    }
                    else if ((_heartbeatTimeout % (100 / _threadWaitTime)) == 0)
                    {
                        _heart = !_heart;
                    }
                    updateState = true;
                }

                if (_rainbowTimeout > 0)
                {
                    _rainbowTimeout--;
                    if (_rainbowTimeout == 0)
                    {
                        _currentColor = _targetColor;
                    }
                    else
                    {
                        _currentColor = (Color)(_rainbowTimeout % 7);
                    }
                    updateState = true;
                }

                //loopCounter++;
                //if (loopCounter > (3000 / _threadWaitTime))
                //{
                //    loopCounter = 0;
                //    updateState = true;
                //}

                if (updateState) {
                    UpdateState();
                }

                Thread.Sleep(_threadWaitTime);
            }
        }

        public void SetColor(Color Color)
        {
            _currentColor = _targetColor = Color;
            UpdateState();
        }

        public void Dance(uint TimeMs=5000)
        {
            _danceTimeout = TimeMs / _threadWaitTime;
        }

        public void FlapWings(uint TimeMs = 5000)
        {
            _flapTimeout = TimeMs / _threadWaitTime;
        }

        public void Heartbeat(uint TimeMs = 5000)
        {
            _heartbeatTimeout = TimeMs / _threadWaitTime;
        }

        public void Rainbow(uint TimeMs = 3000)
        {
            _rainbowTimeout = TimeMs / _threadWaitTime;
        }

        /*
         bit 7:     0 = Heart light on
         bit 4-6:   Head color  0x7 = black (off), 0x6 = red, 0x5 = green, 0x4 = yellow 
                                0x3 = blue, 0x2 = lila, 0x1 = cyan, 0x0 = light blue
         bit 3:     Wings apply force down
         bit 2:     Wings apply force up
                    Wings off: bit 2 = bit 3 = 1
         bit 1:     Turn apply force right
         bit 0:     Turn apply force left
                    Turn off: bit 0 = bit 1 = 1
        */
        private void UpdateState()
        {
            byte state = 0;
            state |= (byte)((int)_currentColor << 4);
            state |= (byte)(_heart ? 0 : 0x80);
            state |= (byte)((int)_flap << 2);
            state |= (byte)(_turn);

            byte[] hidBuffer = new byte[9] {0, 0x55, 0x53, 0x42, 0x43, 0, 0x40, 2, 0};
            if (IsAlive)
            {
                int lpNumberOfBytesWritten = 0;
                hidBuffer[8] = state;
                int result = WriteFile(_hidFileHandle, ref hidBuffer[0], hidBuffer.Length, ref lpNumberOfBytesWritten, 0);
                if (result == 0)
                {
                    LastWin32Error = Marshal.GetLastWin32Error();
                    Trace.WriteLine(string.Format("LyncFellow: WriteFile result=0, GetLastWin32Error()={0}", LastWin32Error));
                    IsAlive = false;
                }
            }
        }

    }
}
