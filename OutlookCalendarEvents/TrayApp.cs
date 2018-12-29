using OutlookCalendarEvents.Properties;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookCalendarEvents
{
    class TrayApp : ApplicationContext
    {
        
        internal RunOptions runOptions = new RunOptions();
        private NotifyIcon trayIcon; 
        private Worker worker;
        private System.Timers.Timer timerPaused; 


        public TrayApp()
        {
            trayIcon = new NotifyIcon()
            {
                Icon = Resources.calendar_32x32Color1, 
                ContextMenu = new ContextMenu(new MenuItem[]
                {
                    new MenuItem("Pause 1 minutes", (sender, e )=>{ Pauser(1); } ), 
                    new MenuItem("Pause 30 minutes", (sender, e )=>{ Pauser(30); } ),
                    new MenuItem("Pause 60 minutes", (sender, e )=>{ Pauser(60); } ),
                    new MenuItem("Resume", (sender, e )=>{ Pauser(0); }  ),
                    new MenuItem("", (sender, e )=>{ } ),                    
                    new MenuItem("Exit", (sender, e)=>{ trayIcon.Visible = false; Application.Exit();  }  )
                })
                , Visible = true 
            };
            EnablePauseDisableResume();
            StartTheWorker(); 


        }


        private void StartTheWorker()
        {
            worker = new Worker();
            worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
            worker.DoWork += worker.DoTheWork;  
            worker.RunWorkerAsync(runOptions); 
        }

        private void Worker_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            Debug.WriteLine("Background Ended"); 
        }

        private void Pauser(int minutesToPause )
        {
            if (minutesToPause == 0 )
            {
                worker._busy.Set(); 
                runOptions.nextCheck = DateTime.MinValue;
                EnablePauseDisableResume();
                if (timerPaused != null )
                {
                    timerPaused.Stop();
                    timerPaused.Dispose();                    
                    timerPaused = null;
                }
            }
            else
            {
                worker._busy.Reset();
                runOptions.nextCheck = DateTime.Now.AddMinutes(minutesToPause);
                DisablePauseEnableResume();
                timerPaused = new System.Timers.Timer();
                timerPaused.Elapsed += (object sender, System.Timers.ElapsedEventArgs e) => { Pauser(0); };
                timerPaused.Interval = minutesToPause * ( 1000 * 60 );
                timerPaused.Start();                 
            }
        }


        private void EnablePauseDisableResume()
        {
            for (int i = 0; i < trayIcon.ContextMenu.MenuItems.Count; i++)
            {
                if (trayIcon.ContextMenu.MenuItems[i].Text.StartsWith("Resume"))
                {
                    trayIcon.ContextMenu.MenuItems[i].Enabled = false;
                }
                if (trayIcon.ContextMenu.MenuItems[i].Text.StartsWith("Pause"))
                {
                    trayIcon.ContextMenu.MenuItems[i].Enabled = true;
                }
            }

        }
        private void DisablePauseEnableResume()
        {
            for (int i = 0; i < trayIcon.ContextMenu.MenuItems.Count; i++)
            {
                if (trayIcon.ContextMenu.MenuItems[i].Text.StartsWith("Resume"))
                {
                    trayIcon.ContextMenu.MenuItems[i].Enabled = true;
                }
                if (trayIcon.ContextMenu.MenuItems[i].Text.StartsWith("Pause"))
                {
                    trayIcon.ContextMenu.MenuItems[i].Enabled = false;
                }
            }
        }
    }
}
