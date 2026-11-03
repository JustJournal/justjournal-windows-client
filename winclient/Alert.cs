using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace JustJournal
{
    partial class Alert
    {
// ReSharper disable InconsistentNaming
        private const Int32 SW_SHOWNOACTIVATE = 4;
// ReSharper restore InconsistentNaming

        // ReSharper disable InconsistentNaming
        private const Int32 SWP_NOACTIVATE = 0x0010;
        private const Int32 HWND_TOPMOST = -1;
        // ReSharper restore InconsistentNaming

        private static int SND_ALIAS = 0x00010000;

        private readonly string _linked;

        public Alert(string Message, string link)
        {
            InitializeComponent();
            msg.Text = Message;
            _linked = link;
            int h = 128;
            int w = 168;
            int cx = Screen.PrimaryScreen.WorkingArea.Width - (w + 25);
            int cy = Screen.PrimaryScreen.WorkingArea.Height - h;
            SetBounds(cx, cy, w, h, BoundsSpecified.All);
            try
            {
                ShowWindow(Handle, SW_SHOWNOACTIVATE);
                SetWindowPos(Handle, HWND_TOPMOST, cx, cy, Width, Height, SWP_NOACTIVATE);
            }
            catch (EntryPointNotFoundException)
            {
                ShowWithoutUser32();
            }
            catch (DllNotFoundException)
            {
                ShowWithoutUser32();
            }
            timer1.Start();
        }

        // Mono outside Windows has no user32, so show it the WinForms way
        private void ShowWithoutUser32()
        {
            TopMost = true;
            Show();
        }

        [DllImport(dllName: "User32.dll")]
        public static extern int ShowWindow(IntPtr hWnd, int cmdShow);

        [DllImport(dllName: "User32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetWindowPos(
            IntPtr hWnd, // handle to window
            Int32 hWndInsertAfter, // placement-order handle
            Int32 X, // horizontal position
            Int32 Y, // vertical position
            Int32 cx, // width
            Int32 cy, // height
            Int32 uFlags // window-positioning options
            );

        [DllImport(dllName: "winmm.dll", CharSet = CharSet.Unicode)]
        private static extern Int32 PlaySound(String lpszName, Int32 hModule, Int32 dwFlags);

        private void Alert_Load(object sender, EventArgs e)
        {
            try
            {
                PlaySound("SystemNotification", 0, SND_ALIAS);
            }
            catch (EntryPointNotFoundException)
            {
                // no winmm outside Windows; skip the sound
            }
            catch (DllNotFoundException)
            {
                // no winmm outside Windows; skip the sound
            }
        }

        private void msg_MouseEnter(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(_linked))
            {
                Font ft = msg.Font;
                msg.ForeColor = Color.Blue;
                msg.Font = new Font(ft.FontFamily, ft.Size, FontStyle.Underline);
                msg.Cursor = Cursors.Hand;
            }
        }

        private void msg_MouseLeave(object sender, EventArgs e)
        {
            Font ft = msg.Font;
            msg.ForeColor = ForeColor;
            msg.Font = new Font(ft.FontFamily, ft.Size, FontStyle.Regular);
            msg.Cursor = Cursors.Default;
        }

        private void msg_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                if (!String.IsNullOrEmpty(_linked))
                {
                    Process.Start(_linked);
                    timer1.Stop();
                    Close();
                }
            }
            catch (Win32Exception xc)
            {
                MessageBox.Show(xc.Message);
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            timer1.Stop();
            Close();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Close();
        }
    }
}