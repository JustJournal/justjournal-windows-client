using System;
using Microsoft.Win32;

namespace JustJournal
{
	/// <summary>
	/// Summary description for Options.
	/// </summary>
	public class Options : System.Windows.Forms.Form
	{
		private System.Windows.Forms.TabControl tabControl1;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.TabPage tabGeneral;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.RadioButton postRaw;
		private System.Windows.Forms.RadioButton postFormatted;
		private System.Windows.Forms.CheckBox chkUseSSL;
		private System.Windows.Forms.CheckBox chkSavePassword;
		private System.Windows.Forms.CheckBox chkAutoLogin;
		private System.Windows.Forms.TabPage tabIntegration;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.CheckBox chkUseWord;
		private System.Windows.Forms.CheckBox chkAutoSpell;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.CheckBox chkMusicDetect;
		private System.Windows.Forms.PictureBox pictureBox1;
		private System.Windows.Forms.CheckBox chkPaused;
		private System.Windows.Forms.CheckBox chkStopped;
		private System.Windows.Forms.GroupBox groupBox5;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Options()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
			var rk = Registry.CurrentUser.CreateSubKey( "SOFTWARE\\JustJournal" );
		    if (rk != null)
		    {
		        RegistryKey mk = rk.CreateSubKey( "Preferences" );

		        chkAutoLogin.Checked = ((string)rk.GetValue( "autologin", "no" )).Equals("yes");
		        chkUseSSL.Checked = ((string)rk.GetValue( "usessl", "yes" )).Equals("yes");
		        string pwd = (string)rk.GetValue( "password", "" );
		        string Format = (string)mk.GetValue( "Formatting", "Formatted" );

		        chkUseWord.Checked = ((string)rk.GetValue( "useword", "no" )).Equals("yes");
		        chkAutoSpell.Checked = ((string)rk.GetValue( "autospell", "no" )).Equals("yes");
		        chkMusicDetect.Checked = ((string)rk.GetValue( "music", "no" )).Equals("yes");
		        chkPaused.Checked = ((string)mk.GetValue( "winampPaused", "no")).Equals("yes");
		        chkStopped.Checked = ((string)mk.GetValue( "winampStopped", "no")).Equals("yes");
		
		        mk.Close();
		        rk.Close();

		        if( pwd.Equals("@") || String.IsNullOrEmpty(pwd) ) 
		            chkSavePassword.Checked = false;
		        else
		            chkSavePassword.Checked = true;

		        if (Format.Equals("Formatted"))
		            postFormatted.Checked = true;
		        else
		            postRaw.Checked = true;
		    }

		    if (!chkMusicDetect.Checked)
            {
                chkPaused.Enabled = false;
                chkStopped.Enabled = false;
            }
		}



		
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			SaveSettings();

			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}	
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Options));
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabGeneral = new System.Windows.Forms.TabPage();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.postRaw = new System.Windows.Forms.RadioButton();
            this.postFormatted = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkUseSSL = new System.Windows.Forms.CheckBox();
            this.chkSavePassword = new System.Windows.Forms.CheckBox();
            this.chkAutoLogin = new System.Windows.Forms.CheckBox();
            this.tabIntegration = new System.Windows.Forms.TabPage();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.chkPaused = new System.Windows.Forms.CheckBox();
            this.chkStopped = new System.Windows.Forms.CheckBox();
            this.chkMusicDetect = new System.Windows.Forms.CheckBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
      
            this.chkAutoSpell = new System.Windows.Forms.CheckBox();
            this.chkUseWord = new System.Windows.Forms.CheckBox();
            this.btnClose = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabGeneral.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabIntegration.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabGeneral);
            this.tabControl1.Controls.Add(this.tabIntegration);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.HotTrack = true;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Multiline = true;
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(432, 406);
            this.tabControl1.TabIndex = 0;
            // 
            // tabGeneral
            // 
            this.tabGeneral.Controls.Add(this.groupBox2);
            this.tabGeneral.Controls.Add(this.groupBox1);
            this.tabGeneral.Location = new System.Drawing.Point(4, 22);
            this.tabGeneral.Name = "tabGeneral";
            this.tabGeneral.Size = new System.Drawing.Size(424, 380);
            this.tabGeneral.TabIndex = 0;
            this.tabGeneral.Text = "General";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.postRaw);
            this.groupBox2.Controls.Add(this.postFormatted);
            this.groupBox2.Location = new System.Drawing.Point(8, 152);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(408, 88);
            this.groupBox2.TabIndex = 2;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Post Mode";
            // 
            // postRaw
            // 
            this.postRaw.Location = new System.Drawing.Point(24, 48);
            this.postRaw.Name = "postRaw";
            this.postRaw.Size = new System.Drawing.Size(104, 24);
            this.postRaw.TabIndex = 1;
            this.postRaw.Text = "Raw Text";
            // 
            // postFormatted
            // 
            this.postFormatted.Location = new System.Drawing.Point(24, 24);
            this.postFormatted.Name = "postFormatted";
            this.postFormatted.Size = new System.Drawing.Size(104, 24);
            this.postFormatted.TabIndex = 0;
            this.postFormatted.Text = "Formatted Text";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chkUseSSL);
            this.groupBox1.Controls.Add(this.chkSavePassword);
            this.groupBox1.Controls.Add(this.chkAutoLogin);
            this.groupBox1.Location = new System.Drawing.Point(8, 16);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(408, 125);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Login";
            // 
            // chkUseSSL
            // 
            this.chkUseSSL.Location = new System.Drawing.Point(24, 88);
            this.chkUseSSL.Name = "chkUseSSL";
            this.chkUseSSL.Size = new System.Drawing.Size(184, 24);
            this.chkUseSSL.TabIndex = 2;
            this.chkUseSSL.Text = "Use secure connection (SSL)";
            // 
            // chkSavePassword
            // 
            this.chkSavePassword.Location = new System.Drawing.Point(24, 56);
            this.chkSavePassword.Name = "chkSavePassword";
            this.chkSavePassword.Size = new System.Drawing.Size(184, 24);
            this.chkSavePassword.TabIndex = 1;
            this.chkSavePassword.Text = "Save password";
            // 
            // chkAutoLogin
            // 
            this.chkAutoLogin.Location = new System.Drawing.Point(24, 24);
            this.chkAutoLogin.Name = "chkAutoLogin";
            this.chkAutoLogin.Size = new System.Drawing.Size(104, 24);
            this.chkAutoLogin.TabIndex = 0;
            this.chkAutoLogin.Text = "Automatic login";
            // 
            // tabIntegration
            // 
            this.tabIntegration.Controls.Add(this.groupBox4);
            this.tabIntegration.Controls.Add(this.groupBox3);
            this.tabIntegration.Location = new System.Drawing.Point(4, 22);
            this.tabIntegration.Name = "tabIntegration";
            this.tabIntegration.Size = new System.Drawing.Size(424, 380);
            this.tabIntegration.TabIndex = 1;
            this.tabIntegration.Text = "Integration";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.groupBox5);
            this.groupBox4.Controls.Add(this.chkMusicDetect);
            this.groupBox4.Controls.Add(this.pictureBox1);
            this.groupBox4.Location = new System.Drawing.Point(8, 144);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(408, 200);
            this.groupBox4.TabIndex = 1;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Music Player";
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.chkPaused);
            this.groupBox5.Controls.Add(this.chkStopped);
            this.groupBox5.Location = new System.Drawing.Point(24, 96);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(240, 88);
            this.groupBox5.TabIndex = 5;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Check Winamp when music is";
            // 
            // chkPaused
            // 
            this.chkPaused.Location = new System.Drawing.Point(24, 24);
            this.chkPaused.Name = "chkPaused";
            this.chkPaused.Size = new System.Drawing.Size(72, 24);
            this.chkPaused.TabIndex = 3;
            this.chkPaused.Text = "Paused";
            // 
            // chkStopped
            // 
            this.chkStopped.Location = new System.Drawing.Point(24, 56);
            this.chkStopped.Name = "chkStopped";
            this.chkStopped.Size = new System.Drawing.Size(72, 24);
            this.chkStopped.TabIndex = 4;
            this.chkStopped.Text = "Stopped";
            // 
            // chkMusicDetect
            // 
            this.chkMusicDetect.Location = new System.Drawing.Point(24, 32);
            this.chkMusicDetect.Name = "chkMusicDetect";
            this.chkMusicDetect.Size = new System.Drawing.Size(240, 24);
            this.chkMusicDetect.TabIndex = 0;
            this.chkMusicDetect.Text = "Detect music";
            this.chkMusicDetect.CheckedChanged += new System.EventHandler(this.chkMusicDetect_CheckedChanged);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(288, 16);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(96, 168);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 2;
            this.pictureBox1.TabStop = false;
            // 
            // groupBox3
            // 

            this.groupBox3.Controls.Add(this.chkAutoSpell);
            this.groupBox3.Controls.Add(this.chkUseWord);
            this.groupBox3.Location = new System.Drawing.Point(8, 16);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(408, 120);
            this.groupBox3.TabIndex = 0;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Microsoft Office";

            // 
            // chkAutoSpell
            // 
            this.chkAutoSpell.Location = new System.Drawing.Point(24, 56);
            this.chkAutoSpell.Name = "chkAutoSpell";
            this.chkAutoSpell.Size = new System.Drawing.Size(256, 24);
            this.chkAutoSpell.TabIndex = 1;
            this.chkAutoSpell.Text = "Automatically check spelling before posting";
            // 
            // chkUseWord
            // 
            this.chkUseWord.Location = new System.Drawing.Point(24, 24);
            this.chkUseWord.Name = "chkUseWord";
            this.chkUseWord.Size = new System.Drawing.Size(240, 24);
            this.chkUseWord.TabIndex = 0;
            this.chkUseWord.Text = "Use Microsoft Word to spell check entries";
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(352, 376);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 1;
            this.btnClose.Text = "Close";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // Options
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(432, 406);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.tabControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Options";
            this.Text = "Options";
            this.TopMost = true;
            this.tabControl1.ResumeLayout(false);
            this.tabGeneral.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.tabIntegration.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox5.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.groupBox3.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		private void btnClose_Click(object sender, EventArgs e)
		{
			SaveSettings();
			Close();
		}

		public void SaveSettings()
		{
			var rk = Registry.CurrentUser.CreateSubKey( "SOFTWARE\\JustJournal" );
		    if (rk != null)
		    {
		        var mk = rk.CreateSubKey( "Preferences" );

		        /*mk.SetValue( "ServerUrl", ServerURL );
			mk.SetValue( "FontFace", FontSetting.Name );
			mk.SetValue( "FontSize", FontSetting.Size );

			mk.SetValue( "ProxyURL", proxyURL.Text );
			mk.SetValue( "ProxyPort", proxyPort.Text );
			mk.SetValue( "ProxyPassword", (useLP.Checked ? proxyPassword.Text : string.Empty) );
			mk.SetValue( "ProxyLogin", (useLP.Checked ? proxyLogin.Text : string.Empty) );

			mk.SetValue( "AutoPlay", (AutoCheckMusic ? "Yes" : "No") );
			mk.SetValue( "StopCheck", (CheckStopped ? "Yes" : "No") );
			mk.SetValue( "PauseCheck", (CheckPaused ? "Yes" : "No") );
			mk.SetValue( "QueuePost", (PostToQueue ? "Yes" : "No") );
			mk.SetValue( "QueueSpeed", QueueSpeed.ToString() );
			
			*/
		        // TODO finish this
		        if (mk != null) mk.SetValue( "Formatting", (postFormatted.Checked ? "Formatted" : "Raw") );
		        rk.SetValue( "music", (chkMusicDetect.Checked ? "yes" : "no"));
		        if (mk != null)
		        {
		            mk.SetValue( "winampPaused", (chkPaused.Checked ? "yes" : "no"));
		            mk.SetValue( "winampStopped", (chkStopped.Checked ? "yes" : "no"));
		        }

		        JustJournalCore.EnableMusicDetection = chkMusicDetect.Checked;
		        JustJournalCore.WinampPaused = chkPaused.Checked;
		        JustJournalCore.WinampStopped = chkStopped.Checked;

		        rk.SetValue("usessl", chkUseSSL.Checked ? "yes" : "no");

		        rk.SetValue("password", chkSavePassword.Checked ? JustJournalCore.Password : "@");

		        rk.SetValue("autologin", chkAutoLogin.Checked ? "yes" : "no");

		        if (chkUseWord.Checked)
		        {
		            JustJournalCore.EnableSpellCheck = true;
		            rk.SetValue( "useword", "yes" );
		        }
		        else
		        {
		            JustJournalCore.EnableSpellCheck = false;
		            rk.SetValue( "useword", "no" );
		        }

		        if (chkUseWord.Checked && chkAutoSpell.Checked)
		        {
		            JustJournalCore.AutoSpellCheck = true;
		            rk.SetValue( "autospell", "yes" );
		        }
		        else
		        {
		            JustJournalCore.AutoSpellCheck = false;
		            rk.SetValue( "autospell", "no" );
		        }

		        /*string old_lang = mk.GetValue( "Language", string.Empty ).ToString();
			string lang = cbLanguage.SelectedItem.ToString();
			foreach( object key in Languages.Keys )
				if( Languages[key].ToString().Equals( lang ) )
					mk.SetValue( "Language", key.ToString() );
				
			if( !old_lang.Equals( mk.GetValue( "Language", string.Empty ).ToString() ) )
				MessageBox.Show( rm.GetString( "message_change_lang" ) );*/

		        if (mk != null) mk.Close();
		    }
		    if (rk != null) rk.Close();
		}

        private void chkMusicDetect_CheckedChanged(object sender, EventArgs e)
        {
            if (!chkMusicDetect.Checked)
            {
                chkPaused.Enabled = false;
                chkStopped.Enabled = false;
            }
            else
            {
                chkPaused.Enabled = true;
                chkStopped.Enabled = true;
            }
        }
	}

}
