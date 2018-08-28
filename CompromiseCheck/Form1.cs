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

        public string allNTLM = "";
        public string compromisedNTLM = "";
        public List<ADObjects> ADO = new List<ADObjects>();

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
            MessageBox.Show(@"If you are prompted to allow changes to the execution policy in the next window, please enter ""Y"".");
            InstallDSinternals();
            ReplicateAD();
            ReadAD();
            checkCompromised();
            listCompromised();
            MessageBox.Show("Process complete!");
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
            if (File.Exists(Environment.CurrentDirectory + @"\adExport.txt"))
            {
                File.Delete(Environment.CurrentDirectory + @"\adExport.txt");
            }
            Process proc = new Process();
            proc.StartInfo.FileName = @"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe";
            proc.StartInfo.Arguments = @"set-executionpolicy remotesigned; Import-Module DSInternals; Get-ADReplAccount -All -NamingContext '" + textBox1.Text + @"' -Server " + textBox2.Text + @" >" + Environment.CurrentDirectory + @"\adExport.txt";
            proc.Start();
            proc.WaitForExit();
        }

        private void ReadAD()
        {
            string objects = File.ReadAllText(Environment.CurrentDirectory + @"\adExport.txt");

            foreach (string record in objects.Split(new string[] { "DistinguishedName: " }, StringSplitOptions.None))
            {
                if (record.Contains(@"SamAccountType: User"))
                {
                    if (checkBox1.Checked)
                    {
                        if (record.Contains(@"Enabled: True"))
                        {
                            addRecord(record);
                        }
                    }else
                    {
                        addRecord(record);
                    }
                }
            }

           File.Delete(Environment.CurrentDirectory + @"\adExport.txt");
        }

        private void addRecord(string record)
        {
            ADObjects obj = new ADObjects();
            foreach (string line in record.Split(new[] {  Environment.NewLine }, StringSplitOptions.None))
            {
                
                if (line.Contains(@"SamAccountName: "))
                {
                    obj.SAM = line.Split(':')[1].TrimStart();
                    //MessageBox.Show(obj.SAM);
                    
                }else
                {
                    if (line.Contains(@"NTHash: "))
                    {
                        string NTLM = line.Split(':')[1].TrimStart().ToUpper();
                        obj.NTLM = NTLM;
                        //MessageBox.Show(obj.NTLM);
                        if (!(allNTLM.Contains(NTLM)))
                        {
                            if (allNTLM != "")
                            {
                                allNTLM += Environment.NewLine;
                            }
                            allNTLM += NTLM;
                            //MessageBox.Show(allNTLM);
                        }
                    }
                }
                
                
            }
            ADO.Add(obj);
        }

        private void checkCompromised()
        {
            string line;
            StreamReader file = new StreamReader(textBox3.Text);
            while ((line = file.ReadLine()) != null)
            {
                string hash = line.Split(':')[0];
                string count = line.Split(':')[1];
                
                if (allNTLM.Contains(hash))
                {
                    if (compromisedNTLM != "")
                    {
                        compromisedNTLM += Environment.NewLine;
                    }
                    compromisedNTLM += hash;
                    allNTLM.Replace(hash, "");
                    
                }
            }
        }

        private void listCompromised()
        {
            foreach (ADObjects obj in ADO)
            {
                if (obj.SAM == "BSemrau")
                {
                    MessageBox.Show(obj.SAM);
                }
                if (compromisedNTLM.Contains(obj.NTLM))
                {
                    if (textBox4.Text != "")
                    {
                        textBox4.Text += Environment.NewLine;
                    }
                    textBox4.Text += obj.SAM;
                }
            }

        }
    }

    public class ADObjects
    {
        public string SAM { get; set; }
        public string NTLM { get; set; }
    }
}
