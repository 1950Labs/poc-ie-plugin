using SHDocVw;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IEExtension
{
    public partial class IE_POC_Form : Form
    {
        public string customUrl { get; set; }
        public IE_POC_Form(string name)
        {
            InitializeComponent();
            customUrl = name;
        }

        private void PictureBox1_Click(object sender, EventArgs e)
        {
            NavigateNewTab("https://www.aa.com/homePage.do?locale=en_US&pref=true");
        }

        private void PictureBox2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void PictureBox3_Click(object sender, EventArgs e)
        {
            NavigateNewTab("https://www.booking.com");
        }

        private void PictureBox4_Click(object sender, EventArgs e)
        {
            NavigateNewTab("https://www.airbnb.com");
        }

        private void PictureBox5_Click(object sender, EventArgs e)
        {
            NavigateNewTab("https://www.latam.com");
        }

        private void PictureBox6_Click(object sender, EventArgs e)
        {
            NavigateNewTab("https://www.copaair.com");
        }

        private void PictureBox7_Click(object sender, EventArgs e)
        {
            NavigateNewTab("https://www.kayak.com");
        }

        private void NavigateNewTab (string url)
        {
            InternetExplorer ie = null;
            ShellWindows allBrowser = new SHDocVw.ShellWindows();
            int browserCount = allBrowser.Count - 1;
            while (browserCount >= 0)
            {
                ie = allBrowser.Item(browserCount) as InternetExplorer;
                if (ie != null && ie.FullName.ToLower().Contains("iexplore.exe"))
                {
                    ie.Navigate2(url, 0x1000);
                    break;
                }
                browserCount--;
            }

            this.Close();
        }
    }
}
