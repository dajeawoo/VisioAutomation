﻿using System.ComponentModel;
using System.Windows.Forms;
using VA=VisioAutomation;

namespace VisioAutomation.UI
{
    public partial class BasicFillControl : UserControl
    {
        public BasicFillControl()
        {
            InitializeComponent();

            this.comboBoxPattern.DataSource = System.Enum.GetValues(typeof(VA.Format.FillPattern));
        }

        [Browsable(true)]
        public System.Drawing.Color ForegroundColor
        {
            get { return this.colorPickerForeground.Color; }
            set { this.colorPickerForeground.Color = value; }
        }

        [Browsable(true)]
        public System.Drawing.Color BackgroundColor
        {
            get { return this.colorPickerBackground.Color; }
            set { this.colorPickerBackground.Color = value; }
        }

        [Browsable(true)]
        public int ForegroundTransparency
        {
            get { return this.ucTransparency1.TransparencyPercent; }
            set { this.ucTransparency1.TransparencyPercent = value; }
        }

        [Browsable(true)]
        public int BackgroundTransparency
        {
            get { return this.ucTransparency2.TransparencyPercent; }
            set { this.ucTransparency2.TransparencyPercent = value; }
        }

        [Browsable(true)]
        public VA.Format.FillPattern FillPattern
        {
            get { return (VA.Format.FillPattern)this.comboBoxPattern.SelectedValue; }
            set { this.comboBoxPattern.SelectedItem = value; }
        }

        private void comboBoxGradient_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            var v = (VA.Format.FillPattern)this.comboBoxPattern.SelectedValue;
        }

        private void toolStripMenuItem1_Click(object sender, System.EventArgs e)
        {
            MessageBox.Show("!");
        }

        private void linkLabelTools_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var form = new FormBasicFillTools();
            form.ForegroundColor = this.ForegroundColor;
            form.BackgroundColor= this.BackgroundColor;

            var results = form.ShowDialog();
            if (results == DialogResult.OK)
            {
                this.ForegroundColor = form.ForegroundColor;
                this.BackgroundColor = form.BackgroundColor;
            }

        }
    }
}