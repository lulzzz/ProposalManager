// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

namespace SmartLink.EncryptTool
{
    partial class Main
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.txtSource = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtTarget = new System.Windows.Forms.TextBox();
            this.btnEncryptString = new System.Windows.Forms.Button();
            this.btnDecryptString = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Source";
            // 
            // txtSource
            // 
            this.txtSource.Location = new System.Drawing.Point(15, 29);
            this.txtSource.Multiline = true;
            this.txtSource.Name = "txtSource";
            this.txtSource.Size = new System.Drawing.Size(514, 65);
            this.txtSource.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 111);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(38, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Target";
            // 
            // txtTarget
            // 
            this.txtTarget.Location = new System.Drawing.Point(15, 127);
            this.txtTarget.Multiline = true;
            this.txtTarget.Name = "txtTarget";
            this.txtTarget.Size = new System.Drawing.Size(514, 65);
            this.txtTarget.TabIndex = 3;
            // 
            // btnEncryptString
            // 
            this.btnEncryptString.Location = new System.Drawing.Point(15, 209);
            this.btnEncryptString.Name = "btnEncryptString";
            this.btnEncryptString.Size = new System.Drawing.Size(86, 23);
            this.btnEncryptString.TabIndex = 4;
            this.btnEncryptString.Text = "Encrypt String";
            this.btnEncryptString.UseVisualStyleBackColor = true;
            this.btnEncryptString.Click += new System.EventHandler(this.btnEncryptString_Click);
            // 
            // btnDecryptString
            // 
            this.btnDecryptString.Location = new System.Drawing.Point(125, 209);
            this.btnDecryptString.Name = "btnDecryptString";
            this.btnDecryptString.Size = new System.Drawing.Size(91, 23);
            this.btnDecryptString.TabIndex = 7;
            this.btnDecryptString.Text = "Decrypt String";
            this.btnDecryptString.UseVisualStyleBackColor = true;
            this.btnDecryptString.Click += new System.EventHandler(this.btnDecryptString_Click);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(547, 249);
            this.Controls.Add(this.btnDecryptString);
            this.Controls.Add(this.btnEncryptString);
            this.Controls.Add(this.txtTarget);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtSource);
            this.Controls.Add(this.label1);
            this.Name = "Main";
            this.Text = "Main";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtSource;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtTarget;
        private System.Windows.Forms.Button btnEncryptString;
        private System.Windows.Forms.Button btnDecryptString;
    }
}