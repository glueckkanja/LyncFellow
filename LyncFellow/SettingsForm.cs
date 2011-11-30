using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Lync.Model;

namespace LyncFellow
{
    public partial class SettingsForm : Form
    {
        public SettingsForm()
        {
            InitializeComponent();

            switch (Properties.Settings.Default.RedOnDndCallBusy)
            {
                case ContactAvailability.DoNotDisturb:
                    OnDoNotDisturb.Checked = true;
                    break;
                case ContactAvailability.Busy:
                    OnBusy.Checked = true;
                    break;
                default:
                    OnCallConference.Checked = true;
                    break;
            }
            DanceOnIncomingCall.Checked = Properties.Settings.Default.DanceOnIncomingCall;
        }

        private void GlueckkanjaLabel_Click(object sender, EventArgs e)
        {
            Process.Start("http://www.glueckkanja.com");
        }

        private void MyOwnWebsiteLabel_Click(object sender, EventArgs e)
        {
            Process.Start("http://lyncfellow.github.com");
        }

        private void CloseButtton_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.DanceOnIncomingCall = DanceOnIncomingCall.Checked;
            if (OnDoNotDisturb.Checked)
            {
                Properties.Settings.Default.RedOnDndCallBusy = ContactAvailability.DoNotDisturb;
            }
            else if (OnBusy.Checked)
            {
                Properties.Settings.Default.RedOnDndCallBusy = ContactAvailability.Busy;
            }
            else
            {
                Properties.Settings.Default.RedOnDndCallBusy = ContactAvailability.None;
            }
            Properties.Settings.Default.Save();
            this.Close();
        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {
            Text += @" v" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        }
    }
}
