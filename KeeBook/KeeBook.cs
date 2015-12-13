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
        static void Main()
        {
        }

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
        private const string KEEBOOK_GROUP_NAME = "Bookmarks (" + ProductName + ")";
        private const string KEEBOOK_LAST_EDITED_NAME = "KeeBook Last Access";

        private const string cbShowSuccessNotificationName = "Show success notification";
        private const string cbShowMsgName = "Show Debug Messages";
        private const string cbDateName = "Set utc date as note";
        private const string cbDuplicateEntry = "Prevent adding duplicate entries";

        static HttpListener listener;
        private Thread httpThread;
        private bool stopped = false;
        private System.ComponentModel.IContainer components;
        private NotifyIcon NotifyIcon1;

        public override bool Initialize(IPluginHost host)
        {
            Terminate();

            if (host == null) return false;
            m_host = host;

            bool showSuccessMsg = Properties.Settings.Default.success_message;
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

            // Add menu item 'Show success message'
            ToolStripMenuItem kbMenuSubItemShowSuccess = new ToolStripMenuItem();
            kbMenuSubItemShowSuccess.Text = cbShowSuccessNotificationName;
            kbMenuSubItemShowSuccess.Checked = showSuccessMsg;
            kbMenuSubItemShowSuccess.Enabled = true;
            kbMenuSubItemShowSuccess.Click += cbShowMessageClick;
            kbMenuTootItem.DropDownItems.Add(kbMenuSubItemShowSuccess);

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
                catch (HttpListenerException ex)
                {
                    showError(ex);
                }
            }
        }

        private void RequestHandler(IAsyncResult r)
        {
            try
            {
                _RequestHandler(r);
            }
            catch (Exception ex)
            {
                showError(ex);
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

            try
            {
                string decoded_url = Encryption.DecryptAESString(url);
                string decoded_title = Encryption.DecryptAESString(title);
                string decoded_icon = Encryption.DecryptAESString(icon);

                showDebugMessage(string.Format("Title - {1}{0}{0} Url - {2}{0}{0} Decoded Title - {3}{0}{0} Decoded Url - {4}{0}{0}", Environment.NewLine, title, url, decoded_title, decoded_url));

                if (creatNewBookmark(decoded_url, decoded_title, decoded_icon))
                {
                    updateLastEdited(decoded_title);
                    customUpdateUI(returnGroup());
                    showSuccess(decoded_title);
                }

            }
            catch (Exception ex)
            {
                showError(ex);
            }
        }

        private bool creatNewBookmark(string url, string title, string icon)
        {
            //If "Set utc date as note" is set in KeeBook, then note = 'the date', else note = string.empty
            string note = (Properties.Settings.Default.date_as_note) ? LongUtcDateIso8601() + " " + ProductName : string.Empty;

            if (!isDuplicateEntry(url, title))
            {
                createEntry(title, url, icon, note);
                return true;
            }

            return false;
        }

        private bool isDuplicateEntry(string url, string title)
        {
            if (Properties.Settings.Default.prevent_duplicate == false)
            {
                return false;
            }

            PwGroup gr = returnGroup();

            foreach (PwEntry entry in gr.Entries)
            {
                try
                {
                    string enTitle = entry.Strings.Get(PwDefs.TitleField).ReadString();
                    string enUrl = entry.Strings.Get(PwDefs.UrlField).ReadString();
                    if (url == enUrl && title == enTitle)
                    {
                        if (Properties.Settings.Default.prevent_duplicate)
                        {
                            showDebugMessage("Not adding duplicate bookmark with title:" + Environment.NewLine + title);
                        }

                        return true;
                    }
                }
                catch (Exception)
                {
                }

            }

            return false;
        }


        private void createEntry(string title, string url = null, string icon = null, string note = null)
        {
            PwGroup pwgroup = returnGroup();
            PwEntry pwe = new PwEntry(true, true);
            setIcon(pwe, icon);
            setTags(pwe, url);
            
            if (!string.IsNullOrEmpty(title))
            {
                pwe.Strings.Set(PwDefs.TitleField, new ProtectedString(false, title));
            }

            if (!string.IsNullOrEmpty(url))
            {
                pwe.Strings.Set(PwDefs.UrlField, new ProtectedString(false, url));
            }

            if (!string.IsNullOrEmpty(note))
            {
                pwe.Strings.Set(PwDefs.NotesField, new ProtectedString(true, note));
            }

            pwgroup.AddEntry(pwe, true);
        }

        private void setTags(PwEntry pwe, string url = null)
        {
            pwe.AddTag("keebook_bookmark");

            if (!string.IsNullOrEmpty(url))
            {
                if (url.Contains("https://"))
                {
                    pwe.AddTag("keebook_secure_url");
                }
            }

            if (pwe.IconId != PwIcon.Star)
            {
                pwe.AddTag("keebook_custom_icon");
            }
        }

        private void setIcon(PwEntry pwe, string icon_url)
        {
            PwCustomIcon custom_icon = getOrAddCustomIcon(getIconBytesFromUrl(icon_url));

            if (custom_icon == null)
                pwe.IconId = PwIcon.Star;
            else
                pwe.CustomIconUuid = custom_icon.Uuid;
        }

        private PwCustomIcon getKeeBookIcon()
        {
            return getOrAddCustomIcon(getIconBytesFromIcon(Properties.Resources.KeeBook));
        }

        private byte[] getIconBytesFromUrl(string icon_url)
        {
            try
            {
                return new System.Net.WebClient().DownloadData(icon_url);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private byte[] getIconBytesFromIcon(Icon icon)
        {
            try
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    icon.Save(ms);
                    return ms.ToArray();
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        private PwCustomIcon getOrAddCustomIcon(byte[] icon_data)
        {
            if (null == icon_data)
            {
                return null;
            }

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

                PwUuid uuid = new PwUuid(true);
                PwCustomIcon custom_icon = new PwCustomIcon(uuid, icon_data);
                m_host.Database.CustomIcons.Add(custom_icon);
                return custom_icon;
            }
            catch (Exception ex)
            {
                showError(ex);
            }

            return null;
        }


        public string CalculateMD5Hash(byte[] input)
        {
            try
            {
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
            catch (Exception)
            {
                return null;
            }
        }

        private void updateLastEdited(string newBookMark)
        {
            var group = returnGroup();
            var oldNote = string.Empty;

            try
            {
                KeePassLib.Collections.PwObjectList<PwEntry> group_entries = group.GetEntries(false);

                for (uint i = 0; i < group_entries.UCount; i++)
                {
                    if (group_entries.GetAt(i).Strings.Get(PwDefs.TitleField).ReadString() == KEEBOOK_LAST_EDITED_NAME)
                    {
                        oldNote = group_entries.GetAt(i).Strings.Get(PwDefs.NotesField).ReadString();
                        group.Entries.RemoveAt(i);
                    }
                }

                PwGroup pwgroup = returnGroup();
                PwEntry pwe = new PwEntry(true, true);
                pwe.CustomIconUuid = getKeeBookIcon().Uuid;
                pwe.Strings.Set(PwDefs.TitleField, new ProtectedString(false, KEEBOOK_LAST_EDITED_NAME));
                pwe.Strings.Set(PwDefs.UserNameField, new ProtectedString(false, LongUtcDateIso8601() + "   (Utc date)"));
                pwe.Strings.Set(PwDefs.NotesField, new ProtectedString(true, oldNote + Environment.NewLine + newBookMark));
                pwgroup.AddEntry(pwe, true);
            }
            catch (Exception ex)
            {
                showError(ex);
            }
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
            m_host.Database.UINeedsIconUpdate = true;

            group = (group == null) ? m_host.Database.RootGroup : group;

            var f = (MethodInvoker)delegate {
                win.UpdateUI(false, null, true, group, true, null, true);
            };

            if (win.InvokeRequired)
                win.Invoke(f);
            else
                f.Invoke();
        }

        public void showSuccess(string title)
        {
            try
            {
                if (Properties.Settings.Default.success_message)
                {
                    NotifyIcon nfi = new NotifyIcon();
                    this.components = new System.ComponentModel.Container();
                    this.NotifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
                    nfi.Visible = true;
                    nfi.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
                    nfi.ShowBalloonTip(1000, "Successfully added item with title :", title, ToolTipIcon.Info);
                }
            }
            catch (Exception ex)
            {
                showError(ex);
            }
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

        private void showError(Exception ex)
        {
            showDebugMessage(String.Format("KeeBook has encountered an error {0}{0} Error Message: {0}{1}{0}{0} Error Trace: {0}{2}{0}{0} Error Source: {0}{3}{0}{0}", Environment.NewLine, ex.Message, ex.StackTrace, ex.Source.ToString()));
        }

        private void cbShowSuccessMessageClick(object sender, EventArgs e)
        {
            Properties.Settings.Default.success_message = !Properties.Settings.Default.success_message;
            ((ToolStripMenuItem)sender).Checked = Properties.Settings.Default.success_message;
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