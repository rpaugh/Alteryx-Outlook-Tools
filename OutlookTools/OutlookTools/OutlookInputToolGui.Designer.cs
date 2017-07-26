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
            this.label3.Location = new System.Drawing.Point(3, 78);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(88, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Folder to Search:";
            // 
            // cboFolderToSearch
            // 
            this.cboFolderToSearch.FormattingEnabled = true;
            this.cboFolderToSearch.Location = new System.Drawing.Point(4, 95);
            this.cboFolderToSearch.Name = "cboFolderToSearch";
            this.cboFolderToSearch.Size = new System.Drawing.Size(315, 21);
            this.cboFolderToSearch.TabIndex = 5;
            // 
            // clbFields
            // 
            this.clbFields.CheckOnClick = true;
            this.clbFields.FormattingEnabled = true;
            this.clbFields.Location = new System.Drawing.Point(4, 139);
            this.clbFields.Name = "clbFields";
            this.clbFields.Size = new System.Drawing.Size(314, 259);
            this.clbFields.TabIndex = 6;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(3, 119);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(40, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Select:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(3, 401);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(109, 13);
            this.label5.TabIndex = 8;
            this.label5.Text = "Save Attachments to:";
            // 
            // lblAttachmentPath
            // 
            this.lblAttachmentPath.AutoSize = true;
            this.lblAttachmentPath.Location = new System.Drawing.Point(119, 401);
            this.lblAttachmentPath.Name = "lblAttachmentPath";
            this.lblAttachmentPath.Size = new System.Drawing.Size(0, 13);
            this.lblAttachmentPath.TabIndex = 9;
            // 
            // btnAttachmentPath
            // 
            this.btnAttachmentPath.Location = new System.Drawing.Point(243, 416);
            this.btnAttachmentPath.Name = "btnAttachmentPath";
            this.btnAttachmentPath.Size = new System.Drawing.Size(75, 23);
            this.btnAttachmentPath.TabIndex = 10;
            this.btnAttachmentPath.Text = "Browse";
            this.btnAttachmentPath.UseVisualStyleBackColor = true;
            this.btnAttachmentPath.Click += new System.EventHandler(this.btnAttachmentPath_Click);
            // 
            // txtAttachmentPath
            // 
            this.txtAttachmentPath.Location = new System.Drawing.Point(4, 418);
            this.txtAttachmentPath.Name = "txtAttachmentPath";
            this.txtAttachmentPath.ReadOnly = true;
            this.txtAttachmentPath.Size = new System.Drawing.Size(233, 20);
            this.txtAttachmentPath.TabIndex = 11;
            // 
            // txtQueryString
            // 
            this.txtQueryString.Location = new System.Drawing.Point(4, 461);
            this.txtQueryString.Name = "txtQueryString";
            this.txtQueryString.Size = new System.Drawing.Size(316, 20);
            this.txtQueryString.TabIndex = 13;
            // 
            // lblQueryStringHelpLink
            // 
            this.lblQueryStringHelpLink.AutoSize = true;
            this.lblQueryStringHelpLink.LinkArea = new System.Windows.Forms.LinkArea(14, 1);
            this.lblQueryStringHelpLink.Location = new System.Drawing.Point(6, 441);
            this.lblQueryStringHelpLink.Name = "lblQueryStringHelpLink";
            this.lblQueryStringHelpLink.Size = new System.Drawing.Size(88, 17);
            this.lblQueryStringHelpLink.TabIndex = 14;
            this.lblQueryStringHelpLink.TabStop = true;
            this.lblQueryStringHelpLink.Text = "Query String (?):";
            this.lblQueryStringHelpLink.UseCompatibleTextRendering = true;
            this.lblQueryStringHelpLink.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lblQueryStringHelpLink_LinkClicked);
            // 
            // OutlookInputToolGui
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
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
            this.Size = new System.Drawing.Size(322, 485);
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
    }
}
