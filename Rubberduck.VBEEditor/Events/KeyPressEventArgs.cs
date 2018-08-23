﻿using System;
using System.Windows.Forms;
using Rubberduck.VBEditor.WindowsApi;

namespace Rubberduck.VBEditor.Events
{
    public class KeyPressEventArgs
    {
        // Note: This offers additional functionality over WindowsApi.KeyPressEventArgs by passing the WndProc arguments.
        public KeyPressEventArgs(IntPtr hwnd, IntPtr wParam, IntPtr lParam, char character = default)
        {
            Hwnd = hwnd;
            WParam = wParam;
            LParam = lParam;
            Character = character;
            if (character == default(char))
            {
                Key = (Keys)wParam;
                if ((User32.GetKeyState(VirtualKeyStates.VK_CONTROL) & 0x8000) != 0)
                {
                    Key |= Keys.Control;
                }
            }
            else
            {
                IsCharacter = true;
            }
        }

        public bool IsCharacter { get; }
        public IntPtr Hwnd { get; }
        public IntPtr WParam { get; }
        public IntPtr LParam { get; }

        public bool Handled { get; set; }

        public char Character { get; }
        public Keys Key { get; }
    }
}
