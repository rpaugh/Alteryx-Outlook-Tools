namespace OutlookTools
{
    partial class OutlookInputToolGui
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.txtUserName = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cboFolderToSearch = new System.Windows.Forms.ComboBox();
            this.clbFields = new System.Windows.Forms.CheckedListBox();
            this.label4 = new System.Windows.Forms.Label();
            this.fbAttachmentPath = new System.Windows.Forms.FolderBrowserDialog();
            this.label5 = new System.Windows.Forms.Label();
            this.lblAttachmentPath = new System.Windows.Forms.Label();
            this.btnAttachmentPath = new System.Windows.Forms.Button();
            this.txtAttachmentPath = new System.Windows.Forms.TextBox();
            this.txtQueryString = new System.Windows.Forms.TextBox();
            this.lblQueryStringHelpLink = new System.Windows.Forms.LinkLabel();
            this.chkIncludeSubFolders = new System.Windows.Forms.CheckBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtSubFolderName = new System.Windows.Forms.TextBox();
            this.chkUseManualServiceURL = new System.Windows.Forms.CheckBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtServiceURL = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "User Name:";
            // 
            // txtUserName
            // 
            this.txtUserName.Location = new System.Drawing.Point(3, 16);
            this.txtUserName.Name = "txtUserName";
            this.txtUserName.Size = new System.Drawing.Size(316, 20);
            this.txtUserName.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 39);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Password:";
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(3, 55);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.PasswordChar = '*';
            this.txtPassword.Size = new System.Drawing.Size(315, 20);
            this.txtPassword.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(3, 124);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(88, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Folder to Search:";
            // 
            // cboFolderToSearch
            // 
            this.cboFolderToSearch.FormattingEnabled = true;
            this.cboFolderToSearch.Location = new System.Drawing.Point(4, 141);
            this.cboFolderToSearch.Name = "cboFolderToSearch";
            this.cboFolderToSearch.Size = new System.Drawing.Size(315, 21);
            this.cboFolderToSearch.TabIndex = 5;
            // 
            // clbFields
            // 
            this.clbFields.CheckOnClick = true;
            this.clbFields.FormattingEnabled = true;
            this.clbFields.Location = new System.Drawing.Point(4, 234);
            this.clbFields.Name = "clbFields";
            this.clbFields.Size = new System.Drawing.Size(314, 259);
            this.clbFields.TabIndex = 6;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(3, 214);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(40, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Select:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(3, 496);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(109, 13);
            this.label5.TabIndex = 8;
            this.label5.Text = "Save Attachments to:";
            // 
            // lblAttachmentPath
            // 
            this.lblAttachmentPath.AutoSize = true;
            this.lblAttachmentPath.Location = new System.Drawing.Point(119, 496);
            this.lblAttachmentPath.Name = "lblAttachmentPath";
            this.lblAttachmentPath.Size = new System.Drawing.Size(0, 13);
            this.lblAttachmentPath.TabIndex = 9;
            // 
            // btnAttachmentPath
            // 
            this.btnAttachmentPath.Location = new System.Drawing.Point(243, 511);
            this.btnAttachmentPath.Name = "btnAttachmentPath";
            this.btnAttachmentPath.Size = new System.Drawing.Size(75, 23);
            this.btnAttachmentPath.TabIndex = 10;
            this.btnAttachmentPath.Text = "Browse";
            this.btnAttachmentPath.UseVisualStyleBackColor = true;
            this.btnAttachmentPath.Click += new System.EventHandler(this.btnAttachmentPath_Click);
            // 
            // txtAttachmentPath
            // 
            this.txtAttachmentPath.Location = new System.Drawing.Point(4, 513);
            this.txtAttachmentPath.Name = "txtAttachmentPath";
            this.txtAttachmentPath.Size = new System.Drawing.Size(233, 20);
            this.txtAttachmentPath.TabIndex = 11;
            // 
            // txtQueryString
            // 
            this.txtQueryString.Location = new System.Drawing.Point(4, 556);
            this.txtQueryString.Name = "txtQueryString";
            this.txtQueryString.Size = new System.Drawing.Size(316, 20);
            this.txtQueryString.TabIndex = 13;
            // 
            // lblQueryStringHelpLink
            // 
            this.lblQueryStringHelpLink.AutoSize = true;
            this.lblQueryStringHelpLink.LinkArea = new System.Windows.Forms.LinkArea(14, 1);
            this.lblQueryStringHelpLink.Location = new System.Drawing.Point(6, 536);
            this.lblQueryStringHelpLink.Name = "lblQueryStringHelpLink";
            this.lblQueryStringHelpLink.Size = new System.Drawing.Size(88, 17);
            this.lblQueryStringHelpLink.TabIndex = 14;
            this.lblQueryStringHelpLink.TabStop = true;
            this.lblQueryStringHelpLink.Text = "Query String (?):";
            this.lblQueryStringHelpLink.UseCompatibleTextRendering = true;
            this.lblQueryStringHelpLink.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lblQueryStringHelpLink_LinkClicked);
            // 
            // chkIncludeSubFolders
            // 
            this.chkIncludeSubFolders.AutoSize = true;
            this.chkIncludeSubFolders.Location = new System.Drawing.Point(6, 169);
            this.chkIncludeSubFolders.Name = "chkIncludeSubFolders";
            this.chkIncludeSubFolders.Size = new System.Drawing.Size(126, 17);
            this.chkIncludeSubFolders.TabIndex = 15;
            this.chkIncludeSubFolders.Text = "Include Sub-Folders?";
            this.chkIncludeSubFolders.UseVisualStyleBackColor = true;
            this.chkIncludeSubFolders.Click += new System.EventHandler(this.chkIncludeSubFolders_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(24, 189);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(92, 13);
            this.label6.TabIndex = 16;
            this.label6.Text = "Sub-Folder Name:";
            // 
            // txtSubFolderName
            // 
            this.txtSubFolderName.Enabled = false;
            this.txtSubFolderName.Location = new System.Drawing.Point(122, 186);
            this.txtSubFolderName.Name = "txtSubFolderName";
            this.txtSubFolderName.Size = new System.Drawing.Size(196, 20);
            this.txtSubFolderName.TabIndex = 17;
            // 
            // chkUseManualServiceURL
            // 
            this.chkUseManualServiceURL.AutoSize = true;
            this.chkUseManualServiceURL.Location = new System.Drawing.Point(6, 81);
            this.chkUseManualServiceURL.Name = "chkUseManualServiceURL";
            this.chkUseManualServiceURL.Size = new System.Drawing.Size(153, 17);
            this.chkUseManualServiceURL.TabIndex = 18;
            this.chkUseManualServiceURL.Text = "Use Manual Service URL?";
            this.chkUseManualServiceURL.UseVisualStyleBackColor = true;
            this.chkUseManualServiceURL.Click += new System.EventHandler(this.chkUseManualServiceURL_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(24, 101);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(71, 13);
            this.label7.TabIndex = 19;
            this.label7.Text = "Service URL:";
            // 
            // txtServiceURL
            // 
            this.txtServiceURL.Enabled = false;
            this.txtServiceURL.Location = new System.Drawing.Point(101, 98);
            this.txtServiceURL.Name = "txtServiceURL";
            this.txtServiceURL.Size = new System.Drawing.Size(217, 20);
            this.txtServiceURL.TabIndex = 20;
            // 
            // OutlookInputToolGui
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.Controls.Add(this.txtServiceURL);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.chkUseManualServiceURL);
            this.Controls.Add(this.txtSubFolderName);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.chkIncludeSubFolders);
            this.Controls.Add(this.lblQueryStringHelpLink);
            this.Controls.Add(this.txtQueryString);
            this.Controls.Add(this.txtAttachmentPath);
            this.Controls.Add(this.btnAttachmentPath);
            this.Controls.Add(this.lblAttachmentPath);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.clbFields);
            this.Controls.Add(this.cboFolderToSearch);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtPassword);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtUserName);
            this.Controls.Add(this.label1);
            this.Name = "OutlookInputToolGui";
            this.Size = new System.Drawing.Size(322, 580);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtUserName;
        private System.Windows.Forms.Label label2;
        public System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cboFolderToSearch;
        private System.Windows.Forms.CheckedListBox clbFields;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.FolderBrowserDialog fbAttachmentPath;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label lblAttachmentPath;
        private System.Windows.Forms.Button btnAttachmentPath;
        private System.Windows.Forms.TextBox txtAttachmentPath;
        private System.Windows.Forms.TextBox txtQueryString;
        private System.Windows.Forms.LinkLabel lblQueryStringHelpLink;
        private System.Windows.Forms.CheckBox chkIncludeSubFolders;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtSubFolderName;
        private System.Windows.Forms.CheckBox chkUseManualServiceURL;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtServiceURL;
    }
}
