using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevComponents.DotNetBar;
using SkypeTheme;
using SKYPE4COMLib;
using AxSKYPE4COMLib;
using System.IO;
using System.Globalization;
using System.Diagnostics;

namespace SkypeApps
{
    public partial class Form1 : Form
    {

        public Skype SkypeAPI = ((Skype)Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("830690FC-BF2F-47A6-AC2D-330BCB402664"))));
        private static readonly DateTime epoch = new DateTime(0x7b2, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
        public Form1()
        {
            InitializeComponent();

            tabControl1.set_SelectedTab(tabControlPanel2.get_TabItem());
            tabControl2.set_SelectedTab(tabControlPanel4.get_TabItem());
            listBox1.SelectedIndex = 0;
            DateTime now = DateTime.Now;
            Years.SelectedIndex = 20;
            Months.SelectedIndex = now.Month - 1;
            Days.SelectedIndex = now.Day - 1;
            Hours.SelectedIndex = now.Hour;
            Minutes.SelectedIndex = now.Minute - 1;
            Secondes.SelectedIndex = now.Second;
        }

        private void All_CheckedChanged(object sender)
        {
            if (All.Checked)
            {
                skypeTextbox1.Enabled = false;
            }
            else
            {
                skypeTextbox1.Enabled = true;
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (FormWindowState.Minimized == base.WindowState)
            {
                notifyIcon1.Visible = true;
                notifyIcon1.ShowBalloonTip(500);
                base.Hide();
            }
            else if (base.WindowState == FormWindowState.Normal)
            {
                notifyIcon1.Visible = false;
            }
        }

        private void GetContactsList()
        {
            foreach (User user in SkypeAPI.Friends)
            {
                ImageList list = new ImageList();
                int num = 0;
                list.Images.Add(Resources.Online);
                list.Images.Add(Resources.Away);
                list.Images.Add(Resources.Busy);
                list.Images.Add(Resources.Invisible);
                list.Images.Add(Resources.Skype);
                if (user.OnlineStatus.ToString() == "olsOnline")
                {
                    num = 0;
                }
                else if (user.OnlineStatus.ToString() == "olsAway")
                {
                    num = 1;
                }
                else if (user.OnlineStatus.ToString() == "olsDoNotDisturb")
                {
                    num = 2;
                }
                else if (user.OnlineStatus.ToString() == "olsOffline")
                {
                    num = 3;
                }
                else
                {
                    num = 3;
                }
                listView1.SmallImageList = list;
                string[] items = new string[] { " " + user.FullName, user.Handle };
                ListViewItem item = new ListViewItem(items) {
                    ImageIndex = num
                };
                listView1.Items.Add(item);
                skypeGroupbox2.Text = "Contacts : " + listView1.Items.Count.ToString();
            }
        }

        private void GetProfilInfos()
        {
            skypeTextbox2.set_Text(SkypeAPI.CurrentUserProfile.FullName);
            skypeTextbox3.set_Text(SkypeAPI.CurrentUserProfile.MoodText);
            skypeTextbox4.set_Text(SkypeAPI.CurrentUserProfile.About);
            skypeTextbox5.set_Text(SkypeAPI.CurrentUserProfile.PhoneMobile);
            skypeTextbox6.set_Text(SkypeAPI.CurrentUserProfile.PhoneHome);
            skypeTextbox7.set_Text(SkypeAPI.CurrentUserProfile.Province);
            skypeTextbox8.set_Text(SkypeAPI.CurrentUserProfile.City);
            skypeTextbox9.set_Text(SkypeAPI.CurrentUserProfile.Homepage);
        }

      

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("http://unknowndevteam.org");
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("http://hackandmodz.net");
        }

        private void LinkSkypeProcess_Click(object sender, EventArgs e)
        {
            try
            {
                SkypeAPI.Attach(5, true);
                try
                {
                    skypeGroupbox1.Enabled = true;
                    skypeGroupbox2.Enabled = true;
                    skypeGroupbox3.Enabled = true;
                    skypeCallButton1.Visible = true;
                    SendMessage.Visible = true;
                    GetContactsList();
                    GetProfilInfos();
                }
                catch
                {
                    MessageBox.Show("Error", "Can't Get your infos...", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                }
            }
            catch
            {
                MessageBox.Show("Error", "Can't Link Process", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            skypeTextbox1.Text = (listView1.FocusedItem.SubItems[1].Text);
            skypeTextbox10.Text = (listView1.FocusedItem.SubItems[1].Text);
            skypeTextbox11.Text = (listView1.FocusedItem.SubItems[1].Text);
        }

        private void MessageFunction()
        {
            if (All.Checked)
            {
                foreach (User user in SkypeAPI.Friends)
                {
                    SkypeAPI.SendMessage(user.Handle, richTextBox1.Text + Environment.NewLine + Environment.NewLine + "Send with Unknauwn SkypeApps Manager");
                }
            }
            else
            {
                SkypeAPI.SendMessage(skypeTextbox1.Text, richTextBox1.Text);
            }
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            base.Show();
            base.WindowState = FormWindowState.Normal;
        }

        private void OpenSkype_Click(object sender, EventArgs e)
        {
            try
            {
                SkypeAPI.Client.Start(false, false);
            }
            catch
            {
                MessageBox.Show("Error, Can't Open Skype.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        private void SaveToClipboard()
        {
            Clipboard.SetText(" ");
            int year = int.Parse(Years.Text);
            int month = int.Parse(Months.Text);
            int day = int.Parse(Days.Text);
            int hour = int.Parse(Hours.Text);
            int minute = int.Parse(Minutes.Text);
            int second = int.Parse(Secondes.Text);
            DateTime time = new DateTime(year, month, day, hour, minute, second);
            string str = skypeTextbox11.Text;
            string text = richTextBox3.Text;
            TimeSpan span = (TimeSpan) (time.ToUniversalTime() - epoch);
            string s = string.Format("<quote author=\"{0}\" timestamp=\"{1}\">{2}</quote>", str, span.TotalSeconds, text);
            IDataObject data = new DataObject();
            data.SetData("System.String", text);
            data.SetData("Text", text);
            data.SetData("UnicodeText", text);
            data.SetData("OEMText", text);
            data.SetData("SkypeMessageFragment", new MemoryStream(Encoding.UTF8.GetBytes(s)));
            data.SetData("Locale", new MemoryStream(BitConverter.GetBytes(CultureInfo.CurrentCulture.LCID)));
            Clipboard.SetDataObject(data);
            MessageBox.Show("Success Copied to Clipboard, Past on Skype.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
        }

        private void SendMessage_Click(object sender, EventArgs e)
        {
            if (skypeTextbox1.Text != string.Empty)
            {
                MessageFunction();
            }
            else
            {
                MessageBox.Show("Please Select a Friend before...", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void skypeButton1_Click(object sender, EventArgs e)
        {
            string text = richTextBox2.Text;
            richTextBox2.Text = "<blink>" + text + "</blink>";
        }

        private void skypeButton2_Click(object sender, EventArgs e)
        {
            string text = richTextBox2.Text;
            richTextBox2.Text = "<center>" + text + "</center>";
        }

        private void skypeButton3_Click(object sender, EventArgs e)
        {
            string text = richTextBox2.Text;
            richTextBox2.Text = "<b>" + text + "</b>";
        }

        private void skypeButton4_Click(object sender, EventArgs e)
        {
            string text = richTextBox2.Text;
            richTextBox2.Text = "<u>" + text + "</u>";
        }

        private void skypeCallButton1_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                SkypeAPI.CurrentUserStatus = TUserStatus.cusOnline;
            }
            else if (radioButton2.Checked)
            {
                SkypeAPI.CurrentUserStatus = TUserStatus.cusAway;
            }
            else if (radioButton3.Checked)
            {
                SkypeAPI.CurrentUserStatus = TUserStatus.cusDoNotDisturb;
            }
            else if (radioButton4.Checked)
            {
                SkypeAPI.CurrentUserStatus = TUserStatus.cusInvisible;
            }
            else if (radioButton5.Checked)
            {
                SkypeAPI.CurrentUserStatus = TUserStatus.cusOffline;
            }
        }

        private void skypeCallButton2_Click(object sender, EventArgs e)
        {
            SkypeAPI.CurrentUserProfile.FullName = skypeTextbox2.Text;
            SkypeAPI.CurrentUserProfile.MoodText = skypeTextbox3.Text;
            SkypeAPI.CurrentUserProfile.About = skypeTextbox4.Text;
            SkypeAPI.CurrentUserProfile.PhoneMobile = skypeTextbox5.Text;
            SkypeAPI.CurrentUserProfile.PhoneHome = skypeTextbox6.Text;
            SkypeAPI.CurrentUserProfile.Province = skypeTextbox7.Text;
            SkypeAPI.CurrentUserProfile.City = skypeTextbox8.Text;
            SkypeAPI.CurrentUserProfile.Homepage = skypeTextbox9.Text;
        }

        private void skypeCallButton4_Click(object sender, EventArgs e)
        {
            SkypeAPI.CurrentUserProfile.MoodText = richTextBox2.Text;
        }

        private void skypeCallButton5_Click(object sender, EventArgs e)
        {
            string[] strArray = new string[] { 
                SmileyArt.BUG(), SmileyArt.BearHug(), SmileyArt.Fuck(), SmileyArt.Dick(), SmileyArt.Boobs(), SmileyArt.Boobs2(), SmileyArt.Boobs3(), SmileyArt.KissYou(), SmileyArt.Apple(), SmileyArt.TeddyBear(), SmileyArt.Rabbit(), SmileyArt.ILoveSex(), SmileyArt.SuperMario(), SmileyArt.Butterfly(), SmileyArt.Alright(), SmileyArt.LOL(), 
                SmileyArt.OKLOVE(), SmileyArt.Drink(), SmileyArt.LOVE(), SmileyArt.JerkOff()
             };
            if (skypeTextbox1.Text != string.Empty)
            {
                if (!skypeCheckbox1.Checked)
                {
                    SkypeAPI.SendMessage(skypeTextbox10.Text, strArray[listBox1.SelectedIndex]);
                }
                else
                {
                    foreach (User user in SkypeAPI.Friends)
                    {
                        SkypeAPI.SendMessage(user.Handle, strArray[listBox1.SelectedIndex] + Environment.NewLine + Environment.NewLine + "Send with Unknauwn SkypeApps");
                    }
                }
            }
            else
            {
                MessageBox.Show("Please Select a Friend before...", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void skypeCallButton6_Click(object sender, EventArgs e)
        {
            SaveToClipboard();
        }

        private void skypeCheckbox1_CheckedChanged(object sender)
        {
            if (skypeCheckbox1.Checked)
            {
                skypeTextbox10.Enabled = false;
            }
            else
            {
                skypeTextbox10.Enabled = true;
            }
        }

        private void skypeTheme1_Click(object sender, EventArgs e)
        {
        }

        private void Spam_CheckedChanged(object sender)
        {
            if (Spam.Checked)
            {
                SpamTimer.Start();
            }
            else
            {
                SpamTimer.Stop();
            }
        }

        private void SpamTimer_Tick(object sender, EventArgs e)
        {
            MessageFunction();
        }

        private void startToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SkypeAPI.PlaceCall(listView1.FocusedItem.SubItems[1].Text, "", "", "");
        }

        private void stopToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SkypeAPI.ActiveCalls[1].Finish();
        }
    }
}

