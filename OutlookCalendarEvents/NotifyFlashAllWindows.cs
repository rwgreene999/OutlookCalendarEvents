using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace OutlookCalendarEvents
{

    static class NotifyFlashAllWindows
    {
        public enum ScreenSide
        {
            random = 0, top = 1, left = 2, right = 3, bottom = 4, all=5
        }
        static Random _random = new Random();
        public static void Flash(ScreenSide side = ScreenSide.random)
        {
            int Offset = 40;
            List<Task> task = new List<Task>();

            if (side == ScreenSide.random)
            {
                side = (ScreenSide)_random.Next(1, 5);
            }

            foreach (var screen in System.Windows.Forms.Screen.AllScreens)
            {
                switch (side)
                {
                    case ScreenSide.random:
                        // cant get here 
                        task.Add(Task.Factory.StartNew(() => { UpdateScreen(screen.Bounds.Top, screen.Bounds.Left, Offset, screen.Bounds.Width); }));
                        break;

                    case ScreenSide.top:
                        task.Add(Task.Factory.StartNew(() => { UpdateScreen(screen.Bounds.Top, screen.Bounds.Left, Offset, screen.Bounds.Width); }));
                        break;
                    case ScreenSide.left:
                        task.Add(Task.Factory.StartNew(() => { UpdateScreen(screen.Bounds.Top, screen.Bounds.Left, screen.Bounds.Height, Offset); }));
                        break;
                    case ScreenSide.right:
                        task.Add(Task.Factory.StartNew(() => { UpdateScreen(screen.Bounds.Top, screen.Bounds.Left + screen.Bounds.Width - Offset, screen.Bounds.Height, Offset); }));
                        break;
                    case ScreenSide.bottom:
                        task.Add(Task.Factory.StartNew(() => { UpdateScreen(screen.Bounds.Bottom - Offset, screen.Bounds.Left, Offset, screen.Bounds.Width); }));
                        break;
                    case ScreenSide.all:
                        task.Add(Task.Factory.StartNew(() => { UpdateScreen(screen.Bounds.Top, screen.Bounds.Left, Offset, screen.Bounds.Width); }));
                        task.Add(Task.Factory.StartNew(() => { UpdateScreen(screen.Bounds.Top, screen.Bounds.Left, screen.Bounds.Height, Offset); }));
                        task.Add(Task.Factory.StartNew(() => { UpdateScreen(screen.Bounds.Top, screen.Bounds.Left + screen.Bounds.Width - Offset, screen.Bounds.Height, Offset); }));
                        task.Add(Task.Factory.StartNew(() => { UpdateScreen(screen.Bounds.Bottom - Offset, screen.Bounds.Left, Offset, screen.Bounds.Width); }));
                        break; 
                    default:
                        break;
                }
            }
            Task.WaitAll(task.ToArray());
        }






        private static void UpdateScreen(int top, int left, int height, int width)
        {

            int msDelay = 13;
            Form form = new Form();
            form.FormBorderStyle = FormBorderStyle.None;
            form.BackColor = Color.FromArgb(0, 255, 0);
            form.SetBounds(left, top, width, height);
            form.TopMost = true;
            form.BringToFront();
            form.Show();
            form.SetBounds(left, top, width, height);
            for (int idx = 0; idx < 254; idx++)
            {
                var col = Color.FromArgb(0, form.BackColor.G - 1, form.BackColor.B + 1);
                form.BackColor = col;
                form.Text = "New Window Detected";
                form.Refresh();
                System.Threading.Thread.Sleep(msDelay);
            }
            System.Diagnostics.Debug.WriteLine("    top:{0}, left:{1}, height:{2}, width:{3}", form.Bounds.Top, form.Bounds.Left, form.Bounds.Height, form.Bounds.Width);
            form.Hide();
            form.Close();
            form.Dispose();
        }
    }
}

