using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Resources;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;
using Timer = System.Windows.Forms.Timer;

[assembly: ComVisible(false)]
[assembly: CLSCompliant(true)]

namespace JustJournal
{
    /// <summary>
    ///     Summary description for Login.
    /// </summary>
    public class Login : Form
    {
        private Timer autoLoginTimer;
        private Button btnLogin;
        private IContainer components;
        private GroupBox groupBox1;
        private Label label1;
        private Label label2;

        private MainMenu mainMenu1;
        private MenuItem menuItem1;
        private MenuItem menuItem10;
        private MenuItem menuItem11;
        private MenuItem menuItem12;
        private MenuItem menuItem13;
        private MenuItem menuItem14;
        private MenuItem menuItem15;
        private MenuItem menuItem16;
        private MenuItem menuItem17;
        private MenuItem menuItem18;
        private MenuItem menuItem19;
        private MenuItem menuItem2;
        private MenuItem menuItem3;
        private MenuItem menuItem4;
        private MenuItem menuItem5;
        private MenuItem menuItem6;
        private MenuItem menuItem7;
        private MenuItem menuItem8;
        private MenuItem menuItem9;
        private MenuItem mnuUseSSL;
        protected NotifyIcon notify;
        private ContextMenu notifyMenu;
        private PictureBox pictureBox1;
        private ResourceManager rm;
        protected TextBox txtPassword;
        protected TextBox txtUserName;

        public Login()
        {
            RegistryKey rk = Registry.CurrentUser.CreateSubKey("SOFTWARE\\JustJournal");
            if (rk == null) return;
            RegistryKey mk = rk.CreateSubKey("Preferences");
            if (mk != null)
            {
                var lang = (string) mk.GetValue("Language");
                if (lang != null)
                {
                    Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture(lang);
                    Thread.CurrentThread.CurrentUICulture = new CultureInfo(lang);
                }
            }

            rm = new ResourceManager(typeof (Login));

            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();

            //
            // TODO: Add any constructor code after InitializeComponent call
            //

            rk.Close();
            if (mk != null) mk.Close();
        }

        private void Login_Load(object sender, EventArgs e)
        {
            RegistryKey rk = Registry.CurrentUser.CreateSubKey("SOFTWARE\\JustJournal");
            if (rk == null) return;
            RegistryKey mk = rk.CreateSubKey("Preferences");
            var uname = (string) rk.GetValue("username", "");
            var pwd = (string) rk.GetValue("password", "");
            bool autoLog = ((string) rk.GetValue("autologin", "no")).Equals("yes");
            bool ssl = ((string) rk.GetValue("usessl", "yes")).Equals("yes");
            JustJournalCore.EnableSpellCheck = ((string) rk.GetValue("useword", "no")).Equals("yes");
            JustJournalCore.AutoSpellCheck = ((string) rk.GetValue("autospell", "no")).Equals("yes");
            JustJournalCore.EnableMusicDetection = ((string) rk.GetValue("music", "no")).Equals("yes");
            if (mk != null)
            {
                JustJournalCore.WinampPaused = ((string) mk.GetValue("winampPaused", "no")).Equals("yes");
                JustJournalCore.WinampStopped = ((string) mk.GetValue("winampStopped", "no")).Equals("yes");
                mk.Close();
            }
            rk.Close();

            txtUserName.Text = uname;
            if (pwd.Equals("@"))
            {
                menuItem4.Checked = false;
                txtPassword.Text = "";
            }
            else
            {
                txtPassword.Text = pwd;
                menuItem4.Checked = true;
            }

            mnuUseSSL.Checked = ssl;
            JustJournalCore.EnableSsl = ssl;

            if (autoLog)
            {
                menuItem5.Checked = true;
                Enabled = false;
                autoLoginTimer.Start();
            }
        }

        private void autologinTimer_Tick(object sender, EventArgs e)
        {
            autoLoginTimer.Stop();
            if (menuItem5.Checked)
            {
                btnLogin_Click(sender, e);
            }
        }

        /// <summary>
        ///     The main entry point for the application.
        /// </summary>
        [STAThread]
        private static void Main()
        {
            Application.Run(new Login());
        }

        /// <summary>
        ///     Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            Enabled = false;
            Cursor = Cursors.WaitCursor;
            JustJournalCore.UserName = txtUserName.Text.Trim();
            JustJournalCore.Password = txtPassword.Text.Trim();

            if (mnuUseSSL.Checked)
            {
                JustJournalCore.EnableSsl = true;
                RegistryKey rk = Registry.CurrentUser.CreateSubKey("SOFTWARE\\JustJournal");
                if (rk != null) rk.SetValue("usessl", "yes");
            }
            else
            {
                JustJournalCore.EnableSsl = false;
                RegistryKey rk = Registry.CurrentUser.CreateSubKey("SOFTWARE\\JustJournal");
                if (rk != null) rk.SetValue("usessl", "no");
            }

            try
            {
                string error;
                if (LogIn(out error))
                {
                    Enabled = true;
                    Cursor = Cursors.Default;
                    Hide();
                    RegistryKey rk = Registry.CurrentUser.CreateSubKey("SOFTWARE\\JustJournal");
                    if (rk != null)
                    {
                        rk.SetValue("username", txtUserName.Text);

                        // save password
                        if (!menuItem4.Checked)
                        {
                            rk.SetValue("password", "@");
                        }
                        else
                        {
                            if ((string) rk.GetValue("password", "") == "@")
                            {
                                DialogResult x =
                                    MessageBox.Show(
                                        @"Password will be saved in the registry in clear text for your windows login.  This could be a security issue.  Are you sure?",
                                        @"Save Password",
                                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                rk.SetValue("password", x.Equals(DialogResult.Yes) ? txtPassword.Text : "@");
                            }
                        }

                        // auto login
                        if (menuItem5.Checked)
                        {
                            rk.SetValue("autologin", "yes");
                        }
                        else
                        {
                            rk.SetValue("autologin", "no");
                        }

                        rk.Close();
                    }
                    notify.Visible = true;
                }
                else
                {
                    if (error != null)
                        MessageBox.Show(error, @"Authentication Error",
                                        MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    Enabled = true;
                    Cursor = Cursors.Default;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, @"Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Enabled = true;
                Cursor = Cursors.Default;
            }
        }

        /// <summary>
        ///     Log in, asking for the second factor code when the account has one.
        /// </summary>
        /// <param name="error">why the login failed, or null when the user cancelled</param>
        private bool LogIn(out string error)
        {
            error = null;
            string status = JustJournalCore.Login();
            if (status == JustJournalCore.LoginOk)
                return true;
            if (status != JustJournalCore.LoginTwoFactorNeeded)
            {
                error = "Unable to login.  Make sure your username and password are correct.";
                return false;
            }

            string prompt = JustJournalCore.TwoFactorMethod == "EMAIL"
                                ? "Enter the code Just Journal e-mailed you, or a recovery code."
                                : "Enter the code from your authenticator app, or a recovery code.";
            string message = prompt;
            bool retry = true;
            while (retry)
            {
                string code = PromptForCode(message);
                if (code == null)
                    return false;

                if (JustJournalCore.LoginTwoFactor(code, out retry) == JustJournalCore.LoginOk)
                    return true;

                message = "That code didn't work. " + prompt;
            }

            error = "Too many wrong codes, or the login expired.  Please log in again.";
            return false;
        }

        /// <returns>the code entered, or null when cancelled</returns>
        private string PromptForCode(string prompt)
        {
            using (var dialog = new Form())
            {
                var label = new Label();
                var code = new TextBox();
                var ok = new Button();
                var cancel = new Button();

                dialog.Text = @"Two-Factor Login";
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.MinimizeBox = false;
                dialog.MaximizeBox = false;
                dialog.ShowInTaskbar = false;
                dialog.ClientSize = new System.Drawing.Size(320, 110);

                label.SetBounds(12, 12, 296, 32);
                label.Text = prompt;
                code.SetBounds(12, 48, 296, 20);
                ok.SetBounds(152, 78, 75, 23);
                ok.Text = @"OK";
                ok.DialogResult = DialogResult.OK;
                cancel.SetBounds(233, 78, 75, 23);
                cancel.Text = @"Cancel";
                cancel.DialogResult = DialogResult.Cancel;

                dialog.Controls.AddRange(new Control[] {label, code, ok, cancel});
                dialog.AcceptButton = ok;
                dialog.CancelButton = cancel;

                return dialog.ShowDialog(this) == DialogResult.OK ? code.Text.Trim() : null;
            }
        }

        private void menuItem2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void menuItem7_Click(object sender, EventArgs e)
        {
            new AboutForm().Show();
        }

        private void menuItem5_Click(object sender, EventArgs e)
        {
            if (menuItem5.Checked)
            {
                menuItem5.Checked = false;
            }
            else
            {
                menuItem5.Checked = true;
            }
        }

        private void menuItem4_Click(object sender, EventArgs e)
        {
            if (menuItem4.Checked)
            {
                menuItem4.Checked = false;
            }
            else
            {
                menuItem4.Checked = true;
            }
        }

        private void menuItem8_Click(object sender, EventArgs e)
        {
            new PostForm().Show();
        }

        private void menuItem11_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void menuItem10_Click(object sender, EventArgs e)
        {
            new AboutForm().Show();
        }

        private void notify_DoubleClick(object sender, EventArgs e)
        {
            new PostForm().Show();
        }

        private void menuItem13_Click(object sender, EventArgs e)
        {
            new Options().Show();
        }

        private void menuItem15_Click(object sender, EventArgs e)
        {
            Process.Start("https://www.justjournal.com/users/" + JustJournalCore.UserName);
        }

        private void menuItem16_Click(object sender, EventArgs e)
        {
            Process.Start("https://www.justjournal.com/users/" + JustJournalCore.UserName + "/friends");
        }

        private void menuItem17_Click(object sender, EventArgs e)
        {
            Process.Start("https://www.justjournal.com/users/" + JustJournalCore.UserName + "/calendar");
        }

        private void menuItem14_Click(object sender, EventArgs e)
        {
            Process.Start("https://www.justjournal.com/");
        }

        private void menuItem19_Click(object sender, EventArgs e)
        {
            Process.Start("https://www.justjournal.com/profile/" + JustJournalCore.UserName);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            var resources = new System.ComponentModel.ComponentResourceManager(typeof (Login));
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.txtUserName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnLogin = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.mainMenu1 = new System.Windows.Forms.MainMenu(this.components);
            this.menuItem1 = new System.Windows.Forms.MenuItem();
            this.menuItem2 = new System.Windows.Forms.MenuItem();
            this.menuItem3 = new System.Windows.Forms.MenuItem();
            this.menuItem5 = new System.Windows.Forms.MenuItem();
            this.menuItem4 = new System.Windows.Forms.MenuItem();
            this.mnuUseSSL = new System.Windows.Forms.MenuItem();
            this.menuItem6 = new System.Windows.Forms.MenuItem();
            this.menuItem7 = new System.Windows.Forms.MenuItem();
            this.autoLoginTimer = new System.Windows.Forms.Timer(this.components);
            this.notify = new System.Windows.Forms.NotifyIcon(this.components);
            this.notifyMenu = new System.Windows.Forms.ContextMenu();
            this.menuItem8 = new System.Windows.Forms.MenuItem();
            this.menuItem9 = new System.Windows.Forms.MenuItem();
            this.menuItem10 = new System.Windows.Forms.MenuItem();
            this.menuItem13 = new System.Windows.Forms.MenuItem();
            this.menuItem18 = new System.Windows.Forms.MenuItem();
            this.menuItem15 = new System.Windows.Forms.MenuItem();
            this.menuItem16 = new System.Windows.Forms.MenuItem();
            this.menuItem17 = new System.Windows.Forms.MenuItem();
            this.menuItem19 = new System.Windows.Forms.MenuItem();
            this.menuItem14 = new System.Windows.Forms.MenuItem();
            this.menuItem12 = new System.Windows.Forms.MenuItem();
            this.menuItem11 = new System.Windows.Forms.MenuItem();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize) (this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // txtPassword
            // 
            resources.ApplyResources(this.txtPassword, "txtPassword");
            this.txtPassword.Name = "txtPassword";
            // 
            // txtUserName
            // 
            this.txtUserName.CharacterCasing = System.Windows.Forms.CharacterCasing.Lower;
            resources.ApplyResources(this.txtUserName, "txtUserName");
            this.txtUserName.Name = "txtUserName";
            // 
            // label1
            // 
            this.label1.ForeColor = System.Drawing.Color.MidnightBlue;
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // label2
            // 
            this.label2.ForeColor = System.Drawing.Color.MidnightBlue;
            resources.ApplyResources(this.label2, "label2");
            this.label2.Name = "label2";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.txtUserName);
            this.groupBox1.Controls.Add(this.txtPassword);
            this.groupBox1.ForeColor = System.Drawing.SystemColors.WindowFrame;
            resources.ApplyResources(this.groupBox1, "groupBox1");
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.TabStop = false;
            // 
            // btnLogin
            // 
            this.btnLogin.BackColor = System.Drawing.Color.MidnightBlue;
            resources.ApplyResources(this.btnLogin, "btnLogin");
            this.btnLogin.ForeColor = System.Drawing.Color.White;
            this.btnLogin.Name = "btnLogin";
            this.btnLogin.UseVisualStyleBackColor = false;
            this.btnLogin.Click += new System.EventHandler(this.btnLogin_Click);
            // 
            // pictureBox1
            // 
            resources.ApplyResources(this.pictureBox1, "pictureBox1");
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.TabStop = false;
            // 
            // mainMenu1
            // 
            this.mainMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[]
                {
                    this.menuItem1,
                    this.menuItem3,
                    this.menuItem6
                });
            // 
            // menuItem1
            // 
            this.menuItem1.Index = 0;
            this.menuItem1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[]
                {
                    this.menuItem2
                });
            resources.ApplyResources(this.menuItem1, "menuItem1");
            // 
            // menuItem2
            // 
            this.menuItem2.Index = 0;
            resources.ApplyResources(this.menuItem2, "menuItem2");
            this.menuItem2.Click += new System.EventHandler(this.menuItem2_Click);
            // 
            // menuItem3
            // 
            this.menuItem3.Index = 1;
            this.menuItem3.MenuItems.AddRange(new System.Windows.Forms.MenuItem[]
                {
                    this.menuItem5,
                    this.menuItem4,
                    this.mnuUseSSL
                });
            resources.ApplyResources(this.menuItem3, "menuItem3");
            // 
            // menuItem5
            // 
            this.menuItem5.Index = 0;
            resources.ApplyResources(this.menuItem5, "menuItem5");
            this.menuItem5.Click += new System.EventHandler(this.menuItem5_Click);
            // 
            // menuItem4
            // 
            this.menuItem4.Index = 1;
            resources.ApplyResources(this.menuItem4, "menuItem4");
            this.menuItem4.Click += new System.EventHandler(this.menuItem4_Click);
            // 
            // mnuUseSSL
            // 
            this.mnuUseSSL.Checked = true;
            this.mnuUseSSL.Index = 2;
            resources.ApplyResources(this.mnuUseSSL, "mnuUseSSL");
            // 
            // menuItem6
            // 
            this.menuItem6.Index = 2;
            this.menuItem6.MenuItems.AddRange(new System.Windows.Forms.MenuItem[]
                {
                    this.menuItem7
                });
            resources.ApplyResources(this.menuItem6, "menuItem6");
            // 
            // menuItem7
            // 
            this.menuItem7.Index = 0;
            resources.ApplyResources(this.menuItem7, "menuItem7");
            this.menuItem7.Click += new System.EventHandler(this.menuItem7_Click);
            // 
            // autoLoginTimer
            // 
            this.autoLoginTimer.Interval = 1500;
            this.autoLoginTimer.Tick += new System.EventHandler(this.autologinTimer_Tick);
            // 
            // notify
            // 
            this.notify.ContextMenu = this.notifyMenu;
            resources.ApplyResources(this.notify, "notify");
            this.notify.DoubleClick += new System.EventHandler(this.notify_DoubleClick);
            // 
            // notifyMenu
            // 
            this.notifyMenu.MenuItems.AddRange(new System.Windows.Forms.MenuItem[]
                {
                    this.menuItem8,
                    this.menuItem9,
                    this.menuItem10,
                    this.menuItem13,
                    this.menuItem18,
                    this.menuItem15,
                    this.menuItem16,
                    this.menuItem17,
                    this.menuItem19,
                    this.menuItem14,
                    this.menuItem12,
                    this.menuItem11
                });
            // 
            // menuItem8
            // 
            this.menuItem8.Index = 0;
            resources.ApplyResources(this.menuItem8, "menuItem8");
            this.menuItem8.Click += new System.EventHandler(this.menuItem8_Click);
            // 
            // menuItem9
            // 
            this.menuItem9.Index = 1;
            resources.ApplyResources(this.menuItem9, "menuItem9");
            // 
            // menuItem10
            // 
            this.menuItem10.Index = 2;
            resources.ApplyResources(this.menuItem10, "menuItem10");
            this.menuItem10.Click += new System.EventHandler(this.menuItem10_Click);
            // 
            // menuItem13
            // 
            this.menuItem13.Index = 3;
            resources.ApplyResources(this.menuItem13, "menuItem13");
            this.menuItem13.Click += new System.EventHandler(this.menuItem13_Click);
            // 
            // menuItem18
            // 
            this.menuItem18.Index = 4;
            resources.ApplyResources(this.menuItem18, "menuItem18");
            // 
            // menuItem15
            // 
            this.menuItem15.Index = 5;
            resources.ApplyResources(this.menuItem15, "menuItem15");
            this.menuItem15.Click += new System.EventHandler(this.menuItem15_Click);
            // 
            // menuItem16
            // 
            this.menuItem16.Index = 6;
            resources.ApplyResources(this.menuItem16, "menuItem16");
            this.menuItem16.Click += new System.EventHandler(this.menuItem16_Click);
            // 
            // menuItem17
            // 
            this.menuItem17.Index = 7;
            resources.ApplyResources(this.menuItem17, "menuItem17");
            this.menuItem17.Click += new System.EventHandler(this.menuItem17_Click);
            // 
            // menuItem19
            // 
            this.menuItem19.Index = 8;
            resources.ApplyResources(this.menuItem19, "menuItem19");
            this.menuItem19.Click += new System.EventHandler(this.menuItem19_Click);
            // 
            // menuItem14
            // 
            this.menuItem14.Index = 9;
            resources.ApplyResources(this.menuItem14, "menuItem14");
            this.menuItem14.Click += new System.EventHandler(this.menuItem14_Click);
            // 
            // menuItem12
            // 
            this.menuItem12.Index = 10;
            resources.ApplyResources(this.menuItem12, "menuItem12");
            // 
            // menuItem11
            // 
            this.menuItem11.Index = 11;
            resources.ApplyResources(this.menuItem11, "menuItem11");
            this.menuItem11.Click += new System.EventHandler(this.menuItem11_Click);
            // 
            // Login
            // 
            this.AcceptButton = this.btnLogin;
            resources.ApplyResources(this, "$this");
            this.BackColor = System.Drawing.SystemColors.Control;
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.btnLogin);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.Menu = this.mainMenu1;
            this.Name = "Login";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Load += new System.EventHandler(this.Login_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize) (this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
        }

        #endregion
    }
}