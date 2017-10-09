using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NWin32;
using System.Runtime.InteropServices;

#pragma warning disable 

namespace wm_copydata
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case Win32.WM_COPYDATA:
                    Win32.CopyDataStruct st = (Win32.CopyDataStruct)Marshal.PtrToStructure(m.LParam, typeof(Win32.CopyDataStruct));
                    string strData = Marshal.PtrToStringAnsi(st.lpData, st.cbData);
                    ParseCommand(strData);
                    break;
                default:
                    base.WndProc(ref m);// let the base class deal with it
                    break;
            }
        }

        private bool SendCmd(int hwnd, string args)
        {

            listBox1.Items.Add( string.Format("sending message '{0}' to {1}",args, hwnd) );

            byte[] bytes;
            Win32.CopyDataStruct cds = new Win32.CopyDataStruct();

            bytes = System.Text.Encoding.ASCII.GetBytes(args + "\x00");

            try
            {
                cds.cbData = bytes.Length;
                cds.lpData = Win32.LocalAlloc(0x40, cds.cbData);
                Marshal.Copy(bytes, 0, cds.lpData, bytes.Length);
                cds.dwData = (IntPtr)3;
                Win32.SendMessage((IntPtr)hwnd, Win32.WM_COPYDATA, IntPtr.Zero, ref cds);
            }
            finally
            {
                cds.Dispose();
            }

            return true;
        }

        private void ParseCommand(string arg)
        {
            listBox1.Items.Add("parsing command: " + arg);
            if (arg.IndexOf(',') > 0)
            {
                int hwnd = 0;
                string[] parts = arg.Split(',');
                if (int.TryParse(parts[0], out hwnd))
                {
                    SendCmd(hwnd, "PONG!");
                }
            } 
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                if (Program.cmdline_args.Length == 1)
                {
                    listBox1.Items.Add("Command line: " + Program.cmdline_args[0]);
                    string command = Program.cmdline_args[0];
                    ParseCommand(command);
                }
            }
            catch (Exception ex) { }

        }


    }
}
