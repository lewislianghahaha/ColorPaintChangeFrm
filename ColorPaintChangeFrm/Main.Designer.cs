﻿namespace ColorPaintChangeFrm
{
    partial class Main
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.MainMenu = new System.Windows.Forms.MenuStrip();
            this.tmclose = new System.Windows.Forms.ToolStripMenuItem();
            this.panel1 = new System.Windows.Forms.Panel();
            this.comgenselect = new System.Windows.Forms.ComboBox();
            this.comselect = new System.Windows.Forms.ComboBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.btnimportWhiteExcel = new System.Windows.Forms.Button();
            this.btnimportemptyexcel = new System.Windows.Forms.Button();
            this.btnopen = new System.Windows.Forms.Button();
            this.MainMenu.SuspendLayout();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // MainMenu
            // 
            this.MainMenu.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.MainMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tmclose});
            this.MainMenu.Location = new System.Drawing.Point(0, 0);
            this.MainMenu.Name = "MainMenu";
            this.MainMenu.Size = new System.Drawing.Size(284, 25);
            this.MainMenu.TabIndex = 1;
            this.MainMenu.Text = "MainMenu";
            // 
            // tmclose
            // 
            this.tmclose.Name = "tmclose";
            this.tmclose.Size = new System.Drawing.Size(44, 21);
            this.tmclose.Text = "关闭";
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.comgenselect);
            this.panel1.Controls.Add(this.comselect);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 25);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(284, 51);
            this.panel1.TabIndex = 2;
            // 
            // comgenselect
            // 
            this.comgenselect.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comgenselect.FormattingEnabled = true;
            this.comgenselect.Location = new System.Drawing.Point(11, 26);
            this.comgenselect.Name = "comgenselect";
            this.comgenselect.Size = new System.Drawing.Size(170, 20);
            this.comgenselect.TabIndex = 4;
            // 
            // comselect
            // 
            this.comselect.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comselect.FormattingEnabled = true;
            this.comselect.Location = new System.Drawing.Point(11, 3);
            this.comselect.Name = "comselect";
            this.comselect.Size = new System.Drawing.Size(260, 20);
            this.comselect.TabIndex = 3;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.btnimportWhiteExcel);
            this.panel2.Controls.Add(this.btnimportemptyexcel);
            this.panel2.Controls.Add(this.btnopen);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 76);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(284, 97);
            this.panel2.TabIndex = 3;
            // 
            // btnimportWhiteExcel
            // 
            this.btnimportWhiteExcel.Location = new System.Drawing.Point(32, 67);
            this.btnimportWhiteExcel.Name = "btnimportWhiteExcel";
            this.btnimportWhiteExcel.Size = new System.Drawing.Size(212, 23);
            this.btnimportWhiteExcel.TabIndex = 2;
            this.btnimportWhiteExcel.Text = "导入控色(增白)剂记录EXCEL";
            this.btnimportWhiteExcel.UseVisualStyleBackColor = true;
            // 
            // btnimportemptyexcel
            // 
            this.btnimportemptyexcel.Location = new System.Drawing.Point(32, 37);
            this.btnimportemptyexcel.Name = "btnimportemptyexcel";
            this.btnimportemptyexcel.Size = new System.Drawing.Size(212, 23);
            this.btnimportemptyexcel.TabIndex = 1;
            this.btnimportemptyexcel.Text = "导入纵向(含空格)EXCEL";
            this.btnimportemptyexcel.UseVisualStyleBackColor = true;
            // 
            // btnopen
            // 
            this.btnopen.Location = new System.Drawing.Point(32, 8);
            this.btnopen.Name = "btnopen";
            this.btnopen.Size = new System.Drawing.Size(212, 23);
            this.btnopen.TabIndex = 0;
            this.btnopen.Text = "导入EXCEL";
            this.btnopen.UseVisualStyleBackColor = true;
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 173);
            this.ControlBox = false;
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.MainMenu);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "Main";
            this.Text = "配方数据转换导出工具";
            this.MainMenu.ResumeLayout(false);
            this.MainMenu.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip MainMenu;
        private System.Windows.Forms.ToolStripMenuItem tmclose;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ComboBox comselect;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button btnopen;
        private System.Windows.Forms.Button btnimportemptyexcel;
        private System.Windows.Forms.ComboBox comgenselect;
        private System.Windows.Forms.Button btnimportWhiteExcel;
    }
}

