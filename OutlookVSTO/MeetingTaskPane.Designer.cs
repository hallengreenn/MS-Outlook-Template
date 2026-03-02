using System.Drawing;
using System.Windows.Forms;

namespace OutlookMeetingAddin
{
    partial class MeetingTaskPane
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
            this.panelHeader = new System.Windows.Forms.Panel();
            this.lblHeader = new System.Windows.Forms.Label();
            this.panelPreview = new System.Windows.Forms.Panel();
            this.lblPreview = new System.Windows.Forms.Label();
            this.btnInsertTemplate = new System.Windows.Forms.Button();
            this.panelInfo = new System.Windows.Forms.Panel();
            this.lblInfo = new System.Windows.Forms.Label();
            this.panelHeader.SuspendLayout();
            this.panelPreview.SuspendLayout();
            this.panelInfo.SuspendLayout();
            this.SuspendLayout();
            //
            // panelHeader
            //
            this.panelHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(120)))), ((int)(((byte)(212)))));
            this.panelHeader.Controls.Add(this.lblHeader);
            this.panelHeader.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelHeader.Location = new System.Drawing.Point(0, 0);
            this.panelHeader.Name = "panelHeader";
            this.panelHeader.Padding = new System.Windows.Forms.Padding(15);
            this.panelHeader.Size = new System.Drawing.Size(350, 60);
            this.panelHeader.TabIndex = 0;
            //
            // lblHeader
            //
            this.lblHeader.AutoSize = true;
            this.lblHeader.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Bold);
            this.lblHeader.ForeColor = System.Drawing.Color.White;
            this.lblHeader.Location = new System.Drawing.Point(15, 18);
            this.lblHeader.Name = "lblHeader";
            this.lblHeader.Size = new System.Drawing.Size(200, 25);
            this.lblHeader.TabIndex = 0;
            this.lblHeader.Text = "📋 Møde Skabelon";
            //
            // panelPreview
            //
            this.panelPreview.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(243)))), ((int)(((byte)(242)))), ((int)(((byte)(241)))));
            this.panelPreview.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelPreview.Controls.Add(this.lblPreview);
            this.panelPreview.Location = new System.Drawing.Point(20, 80);
            this.panelPreview.Name = "panelPreview";
            this.panelPreview.Padding = new System.Windows.Forms.Padding(15);
            this.panelPreview.Size = new System.Drawing.Size(310, 280);
            this.panelPreview.TabIndex = 1;
            //
            // lblPreview
            //
            this.lblPreview.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblPreview.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblPreview.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(96)))), ((int)(((byte)(94)))), ((int)(((byte)(92)))));
            this.lblPreview.Location = new System.Drawing.Point(15, 15);
            this.lblPreview.Name = "lblPreview";
            this.lblPreview.Size = new System.Drawing.Size(278, 248);
            this.lblPreview.TabIndex = 0;
            this.lblPreview.Text = @"📌 Formål med mødet
Kort beskrivelse af, hvorfor vi mødes.

📝 Dagsorden/emner
Liste over punkter der skal drøftes.

👥 Roller og ansvar
Hvem er mødeleder og referat.

✅ Beslutninger og næste skridt
Opsummer beslutninger.

💡 Evt.
Tid til spørgsmål.";
            //
            // btnInsertTemplate
            //
            this.btnInsertTemplate.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(120)))), ((int)(((byte)(212)))));
            this.btnInsertTemplate.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnInsertTemplate.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnInsertTemplate.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold);
            this.btnInsertTemplate.ForeColor = System.Drawing.Color.White;
            this.btnInsertTemplate.Location = new System.Drawing.Point(20, 380);
            this.btnInsertTemplate.Name = "btnInsertTemplate";
            this.btnInsertTemplate.Size = new System.Drawing.Size(310, 50);
            this.btnInsertTemplate.TabIndex = 2;
            this.btnInsertTemplate.Text = "👆 INDSÆT MØDE SKABELON";
            this.btnInsertTemplate.UseVisualStyleBackColor = false;
            this.btnInsertTemplate.Click += new System.EventHandler(this.btnInsertTemplate_Click);
            //
            // panelInfo
            //
            this.panelInfo.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(243)))), ((int)(((byte)(249)))), ((int)(((byte)(253)))));
            this.panelInfo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelInfo.Controls.Add(this.lblInfo);
            this.panelInfo.Location = new System.Drawing.Point(20, 450);
            this.panelInfo.Name = "panelInfo";
            this.panelInfo.Padding = new System.Windows.Forms.Padding(12);
            this.panelInfo.Size = new System.Drawing.Size(310, 100);
            this.panelInfo.TabIndex = 3;
            //
            // lblInfo
            //
            this.lblInfo.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblInfo.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblInfo.Location = new System.Drawing.Point(12, 12);
            this.lblInfo.Name = "lblInfo";
            this.lblInfo.Size = new System.Drawing.Size(284, 74);
            this.lblInfo.TabIndex = 0;
            this.lblInfo.Text = @"💡 Dette vindue vises automatisk når du opretter en ny mødeaftale.

Klik på knappen ovenfor for at indsætte skabelonen i mødet!";
            //
            // MeetingTaskPane
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.Controls.Add(this.panelInfo);
            this.Controls.Add(this.btnInsertTemplate);
            this.Controls.Add(this.panelPreview);
            this.Controls.Add(this.panelHeader);
            this.Name = "MeetingTaskPane";
            this.Size = new System.Drawing.Size(350, 600);
            this.panelHeader.ResumeLayout(false);
            this.panelHeader.PerformLayout();
            this.panelPreview.ResumeLayout(false);
            this.panelInfo.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.Panel panelHeader;
        private System.Windows.Forms.Label lblHeader;
        private System.Windows.Forms.Panel panelPreview;
        private System.Windows.Forms.Label lblPreview;
        private System.Windows.Forms.Button btnInsertTemplate;
        private System.Windows.Forms.Panel panelInfo;
        private System.Windows.Forms.Label lblInfo;
    }
}
