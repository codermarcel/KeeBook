using System;
using System.Windows.Forms;
using System.Net;
using System.Threading;

using KeePass.Forms;
using KeePass.Plugins;
using KeePassLib;
using KeePassLib.Security;
using System.Drawing;

namespace KeeBook
{
    public sealed class KeeBookExt : Plugin
    {

        #region Date and Time

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


        string lastUri = string.Empty;

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

            // Add a separator at the bottom
            ToolStripSeparator kbSeperator = new ToolStripSeparator();
            tsMenu.Add(kbSeperator);

            // Add the popup menu item
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

            lastUri = req.Url.ToString();
            string title = req.QueryString["t"];
            string url = req.QueryString["u"];
            string icon = req.QueryString["i"];

            string decoded_title = System.Uri.UnescapeDataString(title);
            string decoded_url = System.Uri.UnescapeDataString(url);

            if (Properties.Settings.Default.debug_message) 
            {
                MessageBox.Show("Title :" + title + Environment.NewLine + "Url : " + url + "Decoded Title :" + decoded_title + Environment.NewLine + "Decoded Url : " + decoded_url, ProductName + " Debug Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            creatNewBookmark(decoded_title, decoded_url, icon, Properties.Settings.Default.date_as_note, Properties.Settings.Default.prevent_duplicate);
        }

        private void creatNewBookmark(string title, string url, string icon, Boolean addUtcDate = true, Boolean checkDuplicate = true)
        {
            if (checkDuplicate)
            {
                if (isDuplicateEntry(title, url))
                {
                    if (Properties.Settings.Default.debug_message)
                    {
                        MessageBox.Show("Not adding Duplicate with title" + Environment.NewLine + title, ProductName + " Debug Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    return;
                }
            }

            MessageBox.Show(icon);



            PwGroup pwgroup = returnGroup();
            PwEntry pwe = new PwEntry(true, true);
            pwe.IconId = getEntryIcon(icon);
            pwe.Strings.Set(PwDefs.TitleField, new ProtectedString(false, title));
            pwe.Strings.Set(PwDefs.UrlField, new ProtectedString(false, url));



            if (addUtcDate)
            {
                pwe.Strings.Set(PwDefs.NotesField, new ProtectedString(true, LongUtcDateIso8601() + Environment.NewLine + ProductName));
            }

            pwgroup.AddEntry(pwe, true);
            UpdateUI(returnGroup());
        }

        private bool isDuplicateEntry(string title, string url)
        {
            PwGroup gr = returnGroup();
            foreach (PwEntry entry in gr.Entries)
            {
                string enTitle = entry.Strings.Get(PwDefs.TitleField).ReadString();
                string enUrl = entry.Strings.Get(PwDefs.UrlField).ReadString();
                if (url == enUrl && title == enTitle)
                {
                    return true;
                }
            }
            return false;
        }

        private PwIcon getEntryIcon(string icon)
        {
            return getSecureIcon("https://");
        }

        private PwIcon getSecureIcon(string url)
        {
            if (url.Contains("https://"))
            {
                return PwIcon.LockOpen;
            }
            return PwIcon.Star;
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


        private PwGroup returnGroup()
        {
            var root = m_host.Database.RootGroup;
            var group = root.FindCreateGroup(KEEBOOK_GROUP_NAME, false);
            if (group == null)
            {
                group = new PwGroup(true, true, KEEBOOK_GROUP_NAME, PwIcon.Star);
                root.AddGroup(group, true);
                UpdateUI(null);
            }
            return group;
        }


        private void UpdateUI(PwGroup group)
        {
            var win = m_host.MainWindow;
            if (group == null) group = m_host.Database.RootGroup;
            var f = (MethodInvoker)delegate {
                win.UpdateUI(false, null, true, group, true, null, true);
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
            saveSettings();

            UpdateUI(null);
            m_host = null;
            Environment.Exit(1337);
        }

        public void saveSettings()
        {
            Properties.Settings.Default.Save();
        }
    }
}

