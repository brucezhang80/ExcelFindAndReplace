﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFindAndReplace {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
            InitStuff();
            
        }

        public void InitStuff() {
            openFileDialog1.Filter = "Excel XLSX (*.xlsx)|*.xlsx|" + "Excel All (*.xlsx;*.xls)|*.xlsx;*.xls";
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
        }

        private void toolStripButton1_Click(object sender, EventArgs e) {
            AddFiles();
        }

        private void toolStripLabel1_Click(object sender, EventArgs e) {
            AddFiles();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e) {
            AddFiles();
        }

        //Add files to ListView
        public void AddFiles() {
            if (openFileDialog1.ShowDialog() == DialogResult.OK) {
                foreach (string file in openFileDialog1.FileNames) {
                    ListViewItem item = new ListViewItem(new[] { file });
                    listView1.Items.Add(item);
                }
                toolStripStatusLabel1.Text = "Files added: " + openFileDialog1.FileNames.Length.ToString();
            }
        }

        //Press Del to remove selected items
        private void listView1_KeyDown(object sender, KeyEventArgs e) {
            if (Keys.Delete == e.KeyCode) {
                statusLabel_Del();
                foreach (ListViewItem item in listView1.SelectedItems) {
                    listView1.Items.Remove(item);
                }  
            }
        }

        //Press Remove button in toolstrip to remove items
        private void toolStripLabel2_Click(object sender, EventArgs e) {
            statusLabel_Del();
            foreach (ListViewItem item in listView1.SelectedItems) {
                listView1.Items.Remove(item);
            }
        }

        public void statusLabel_Del() => toolStripStatusLabel1.Text = "Files removed: " + listView1.SelectedItems.Count.ToString();
    }
}
