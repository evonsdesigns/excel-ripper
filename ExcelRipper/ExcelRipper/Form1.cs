//
// EvonsDesigns Excel-Ripper Application for Windows
// The following is copyright 2012 EvonsDesigns
// Author: Joe Evans (evonsdesigns@gmail.com)
//
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Windows.Forms;

// Class for the Form
namespace ExcelRipper
{
    public partial class Form1 : Form
    {
        // Instance of the excel manager for making calls back to it.
        private readonly ExcelManager manager;

        /// <summary>
        /// Constructor for the form. Creates an instance of the ExcelManager class.
        /// </summary>
        public Form1()
        {
            InitializeComponent();
            manager = new ExcelManager(this); // takes the given form to manipulate
        }

        /// <summary>
        /// Save button logic
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveButton_Click(object sender, EventArgs e)
        {
            DialogResult result = saveFileDialog1.ShowDialog();
            if(result == DialogResult.Yes)
            {
                
            }
        }

        /// <summary>
        /// Select files button logic
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void loadButton_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.Yes)
            {

            }
        }

        /// <summary>
        /// Sets the given save file for the user
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            tb_outfile.Text = saveFileDialog1.FileName;
        }

        /// <summary>
        /// Appends all of the given filenames selected to the files for ripping.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            foreach(string s in openFileDialog1.FileNames)
            {
                tb_files.Text += s + Environment.NewLine;
            }

            openFileDialog1.Reset();

        }

        /// <summary>
        /// Rip It button - starts a new task for the ExcelManager
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DoWorkSonButton_Click(object sender, EventArgs e)
        {
            List<string> items = new List<string>();
            items.AddRange(tb_files.Lines);

            Task.Factory.StartNew(() => manager.Work(items, tb_range.Text, tb_outfile.Text, tb_columnheader.Text));

        }

        /// <summary>
        /// Method to allow the updating of the status to show for the user.
        /// </summary>
        /// <param name="text"></param>
        public void SetStatus(string text)
        {
            statusLabel.Text = text;
        }
    }
}
