using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;

namespace CompromiseCheck
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start(@"https://haveibeenpwned.com/Passwords");
            linkLabel1.LinkVisited = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            textBox3.Text = openFileDialog1.FileName;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Please note, this process may take anywhere from a few minutes to a few days depending on a number of factors.  During this time, the program will appear to be frozen.");
            InstallDSinternals();
            ReplicateAD();
        }
        private void InstallDSinternals()
        {
            if (!(Directory.Exists(@"C:\Program Files\WindowsPowerShell\Modules\DSInternals")))
            {
                try
                {
                    new Microsoft.VisualBasic.Devices.Computer().FileSystem.CopyDirectory(Environment.CurrentDirectory + @"\DSInternals", @"C:\Program Files\WindowsPowerShell\Modules\DSInternals");
                }catch
                {
                    MessageBox.Show(@"Error installing DSInternals modules.  Please install to C:\Program Files\WindowsPowerShell\Modules");
                    Application.Exit();
                }
            }
        }
        private void ReplicateAD()
        {
            Process proc = new Process();
            proc.StartInfo.FileName = @"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe";
            proc.StartInfo.Arguments = @"set-executionpolicy remotesigned; Import-Module DSInternals; Get-ADReplAccount -All -NamingContext '" + textBox1.Text + @"' -Server " + textBox2.Text + @" >" + Environment.CurrentDirectory + @"\adExport.txt";
            proc.Start();
            proc.WaitForExit();
        }
    }
}
