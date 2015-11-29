using System;
using System.Windows.Forms;
using System.Net;
using System.Threading;

using KeePass.Forms;
using KeePass.Plugins;
using KeePassLib;
using KeePassLib.Security;
using System.Drawing;
using System.Text;
using System.Security.Cryptography;
using System.IO;

namespace KeeBook
{
    public sealed class KeeBookExt : Plugin
    {

        #region Date and time
        public readonly System.DateTime EPOCH = new System.DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);

        public int UnixTime(System.DateTime d)
        {
            return Convert.ToInt32((d.ToUniversalTime() - EPOCH).TotalSeconds);
        }

        public string ShortUtcDateIso8601()
        {
            return EPOCH.AddSeconds(UnixTime(System.DateTime.UtcNow)).ToUniversalTime().ToString("yyyy.MM.dd");  //OR d  (2015.08.30)
        }

        public string LongUtcDateIso8601()
        {
            return EPOCH.AddSeconds(UnixTime(System.DateTime.UtcNow)).ToUniversalTime().ToString("yyyy.MM.dd - hh:mm");  //OR d  (2015.08.30)
        }
        #endregion

        private IPluginHost m_host = null;
        private const string ProductName = "KeeBook";
        private const string keepass_prefix = ProductName + "_";
        private const string KEEBOOK_GROUP_NAME = "Bookmarks (" + ProductName + ")";

        private const string cbShowMsgName = "Show Debug Messages";
        private const string cvShowMsg = keepass_prefix + cbShowMsgName;
        private const string cbDateName = "Set utc date as note";
        private const string cvDate = keepass_prefix + cbDateName;
        private const string cbDuplicateEntry = "Prevent adding duplicate entries";
        private const string cvDuplicateEntry = keepass_prefix + cbDuplicateEntry;

        static HttpListener listener;
        private Thread httpThread;
        private bool stopped = false;

        public override bool Initialize(IPluginHost host)
        {
            Terminate();

            if (host == null) return false;
            m_host = host;

            bool showMsg = Properties.Settings.Default.debug_message;
            bool setDate = Properties.Settings.Default.date_as_note;
            bool checkDuplicate = Properties.Settings.Default.prevent_duplicate;

            ToolStripItemCollection tsMenu = m_host.MainWindow.ToolsMenu.DropDownItems;

            // Add a separator at the bottom of KeePass
            ToolStripSeparator kbSeperator = new ToolStripSeparator();
            tsMenu.Add(kbSeperator);

            // Add the popup menu item to KeePass
            ToolStripMenuItem kbMenuTootItem = new ToolStripMenuItem();
            kbMenuTootItem.Text = ProductName;
            kbMenuTootItem.Image = KeeBook.Properties.Resources.KeeBook.ToBitmap();
            tsMenu.Add(kbMenuTootItem);

            // Add menu item 'Show entry messagebox'
            ToolStripMenuItem kbMenuSubItemShowmsg = new ToolStripMenuItem();
            kbMenuSubItemShowmsg.Text = cbShowMsgName;
            kbMenuSubItemShowmsg.Checked = showMsg;
            kbMenuSubItemShowmsg.Enabled = true;
            kbMenuSubItemShowmsg.Click += cbShowMessageClick;
            kbMenuTootItem.DropDownItems.Add(kbMenuSubItemShowmsg);

            // Add submenu item "Set date as note"
            ToolStripMenuItem kbMenuSubItemCbDate = new ToolStripMenuItem();
            kbMenuSubItemCbDate.Text = cbDateName;
            kbMenuSubItemCbDate.Checked = setDate;
            kbMenuSubItemCbDate.Enabled = true;
            kbMenuSubItemCbDate.Click += cbSetDateClick;
            kbMenuTootItem.DropDownItems.Add(kbMenuSubItemCbDate);

            // Add submenu item "Check for duplicate entry"
            ToolStripMenuItem kbPreventDuplicates = new ToolStripMenuItem();
            kbPreventDuplicates.Text = cbDuplicateEntry;
            kbPreventDuplicates.Checked = checkDuplicate;
            kbPreventDuplicates.Enabled = true;
            kbPreventDuplicates.Click += cbPreventDuplicateClick;
            kbMenuTootItem.DropDownItems.Add(kbPreventDuplicates);

            reqListener();
            return true;
        }

        private void reqListener()
        {
            listener = new HttpListener();
            listener.Prefixes.Add("http://localhost:1339/");
            listener.AuthenticationSchemes = AuthenticationSchemes.Anonymous;
            listener.Start();
            httpThread = new Thread(new ThreadStart(Run));
            httpThread.Start();
        }

        private void Run()
        {
            while (!stopped)
            {
                try
                {
                    var r = listener.BeginGetContext(new AsyncCallback(RequestHandler), listener);
                    r.AsyncWaitHandle.WaitOne();
                    r.AsyncWaitHandle.Close();
                }
                catch (ThreadInterruptedException) { }
                catch (HttpListenerException e)
                {
                    MessageBox.Show(ProductName + " error (1)" + Environment.NewLine + e.ToString());
                }
            }
        }

        private void RequestHandler(IAsyncResult r)
        {
            try
            {
                _RequestHandler(r);
            }
            catch (Exception e)
            {
                MessageBox.Show(ProductName + " error (2)" + Environment.NewLine + e.ToString());
            }
        }

        private void _RequestHandler(IAsyncResult r)
        {
            if (stopped) return;

            var l = (HttpListener)r.AsyncState;
            var ctx = l.EndGetContext(r);
            var req = ctx.Request;
            var resp = ctx.Response;

            string lastUri = req.Url.ToString();
            string url = req.QueryString["u"];
            string title = req.QueryString["t"];
            string icon = req.QueryString["i"];

            string decoded_url = url;
            string decoded_title = title;
            string decoded_icon = url;

            try
            {
                decoded_url = System.Text.UTF8Encoding.UTF8.GetString(Convert.FromBase64String(url));
                decoded_title = System.Text.UTF8Encoding.UTF8.GetString(Convert.FromBase64String(title));
                decoded_icon = System.Text.UTF8Encoding.UTF8.GetString(Convert.FromBase64String(icon));

                //Show debug message
                showDebugMessage(string.Format("Title - {1}{0}{0} Url - {2}{0}{0} Decoded Title - {3}{0}{0} Decoded Url - {4}{0}{0}", Environment.NewLine, title, url, decoded_title, decoded_url));

                //Check if entry already exists.
                if (!isDuplicateEntry(decoded_url, decoded_title))
                {
                    creatNewBookmark(decoded_url, decoded_title, decoded_icon);
                }
            }
            catch (Exception ex)
            {
                showDebugMessage(ex.Message);
            }
        }

        private bool isDuplicateEntry(string url, string title)
        {
            PwGroup gr = returnGroup();

            foreach (PwEntry entry in gr.Entries)
            {
                string enTitle = entry.Strings.Get(PwDefs.TitleField).ReadString();
                string enUrl = entry.Strings.Get(PwDefs.UrlField).ReadString();
                if (url == enUrl && title == enTitle)
                {
                    if (Properties.Settings.Default.prevent_duplicate)
                    {
                        showDebugMessage("Not adding Duplicate with title" + Environment.NewLine + title);
                    }

                    return true;
                }
            }

            return false;
        }


        private void creatNewBookmark(string url, string title, string icon)
        {
            PwGroup pwgroup = returnGroup();
            PwEntry pwe = new PwEntry(true, true);

            PwCustomIcon custom_icon = addCustomIcon(icon);

            if (custom_icon != null)
            {
                pwe.CustomIconUuid = custom_icon.Uuid;
            }else
            {
                pwe.IconId = PwIcon.Star;
            }

            pwe.Strings.Set(PwDefs.TitleField, new ProtectedString(false, title));
            pwe.Strings.Set(PwDefs.UrlField, new ProtectedString(false, url));

            if (Properties.Settings.Default.date_as_note)
            {
                pwe.Strings.Set(PwDefs.NotesField, new ProtectedString(true, LongUtcDateIso8601() + Environment.NewLine + ProductName));
            }

            pwgroup.AddEntry(pwe, true);
            customUpdateUI(returnGroup());
        }


        private PwCustomIcon addCustomIcon(string icon_url)
        {

            try
            {
                byte[] icon_data = new System.Net.WebClient().DownloadData(icon_url);

                PwCustomIcon check_icon = getCustomIcon(icon_data);
                
                if (check_icon != null)
                {
                    return check_icon;
                }


                PwUuid uuid = new PwUuid(true);
                PwCustomIcon custom_icon = new PwCustomIcon(uuid, icon_data);
                m_host.Database.CustomIcons.Add(custom_icon);
                return custom_icon;
            }
            catch (Exception ex)
            {
                showDebugMessage(ex.Message);
            }

            return null;
        }

        private PwCustomIcon getCustomIcon(byte[] icon_data)
        {
            try
            {
                System.Collections.Generic.List<PwCustomIcon> icon_entries = m_host.Database.CustomIcons;

                for (int i = 0; i < icon_entries.Count; i++)
                {
                    string checksum1 = CalculateMD5Hash(icon_entries[i].ImageDataPng);
                    string checksum2 = CalculateMD5Hash(icon_data);

                    if (checksum1 == checksum2)
                    {
                        return icon_entries[i];
                    }
                }

            }
            catch (Exception ex)
            {
                showDebugMessage(ex.Message);
            }

            return null;
        }

        public string CalculateMD5Hash(byte[] input)
        {
            // step 1, calculate MD5 hash from input
            System.Security.Cryptography.MD5 md5 = System.Security.Cryptography.MD5.Create();
            byte[] hash = md5.ComputeHash(input);

            // step 2, convert byte array to hex string
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < hash.Length; i++)
            {
                sb.Append(hash[i].ToString("X2"));
            }
            return sb.ToString();
        }


        private PwGroup returnGroup()
        {
            var root = m_host.Database.RootGroup;
            var group = root.FindCreateGroup(KEEBOOK_GROUP_NAME, false);

            if (group == null)
            {
                group = new PwGroup(true, true, KEEBOOK_GROUP_NAME, PwIcon.Star);
                root.AddGroup(group, true);
                customUpdateUI(null);
            }

            return group;
        }

        private void customUpdateUI(PwGroup group)
        {
            var win = m_host.MainWindow;
            if (group == null) group = m_host.Database.RootGroup;
            var f = (MethodInvoker)delegate {
                win.UpdateUI(true, null, true, group, true, null, true);
            };
            if (win.InvokeRequired)
                win.Invoke(f);
            else
                f.Invoke();
        }

        public override void Terminate()
        {
            if (m_host == null) return;
            stopped = true;
            listener.Stop();
            listener.Close();
            httpThread.Interrupt();
            listener = new HttpListener();
            httpThread = null;

            if (Properties.Settings.Default.debug_message)
            {
                MessageBox.Show("Saving settings");
            }

            Properties.Settings.Default.Save();
        
            customUpdateUI(null);
            m_host = null;
            Environment.Exit(1337);
        }

        private void showDebugMessage(string message)
        {
            if (Properties.Settings.Default.debug_message)
            {
                MessageBox.Show(message, ProductName + " Debug Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void cbShowMessageClick(object sender, EventArgs e)
        {
            Properties.Settings.Default.debug_message = !Properties.Settings.Default.debug_message;
            ((ToolStripMenuItem)sender).Checked = Properties.Settings.Default.debug_message;
        }

        private void cbSetDateClick(object sender, EventArgs e)
        {
            Properties.Settings.Default.date_as_note = !Properties.Settings.Default.date_as_note;
            ((ToolStripMenuItem)sender).Checked = Properties.Settings.Default.date_as_note;
        }

        private void cbPreventDuplicateClick(object sender, EventArgs e)
        {
            Properties.Settings.Default.prevent_duplicate = !Properties.Settings.Default.prevent_duplicate;
            ((ToolStripMenuItem)sender).Checked = Properties.Settings.Default.prevent_duplicate;
        }

    }
}