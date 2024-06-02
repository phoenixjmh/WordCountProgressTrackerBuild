﻿using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using System;
using System.Drawing;
using System.Windows.Controls.Primitives;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordCountProgressTracker
{
    public partial class WordCountProgressTracker
    {
        //=== Set Form References ===
        private ProgressBar progressBar;
        private Form statusBarForm;
        private Label label;
        private Button changeCountButton;
        private int goalWordCount = 0;
        private bool init=true;

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            PromptForWordCountGoal();

            InitializeStatusBar();

            this.Application.DocumentBeforeSave += ThisAddIn_DocumentBeforeSave; // Update on document change
            this.Application.DocumentChange += ThisAddIn_DocumentChange; // Update on document change
            this.Application.WindowActivate += Application_WindowActivate; // Update on window activation
            init = false;
        }

        private void ThisAddIn_DocumentBeforeSave(Document doc, ref bool SaveAsUI, ref bool Cancel)
        {
            int currentWordCount = doc.Words.Count;
            double progressPercentage = (double)currentWordCount / goalWordCount * 100;
            UpdateStatusBar(currentWordCount, progressPercentage);
        }
        private void ThisAddIn_DocumentChange()
        {
            Document doc = this.Application.ActiveDocument;
            progressBar.Maximum = goalWordCount;
            int currentWordCount = doc.Words.Count;
            double progressPercentage = (double)currentWordCount / goalWordCount * 100;
            UpdateStatusBar(currentWordCount, progressPercentage);
        }
        private void PromptForWordCountGoal()
        {
            // Create a new form for input
            Form prompt = new Form()
            {
                Width = 500,
                Height = 150,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "Word Count Goal",
                StartPosition = FormStartPosition.CenterScreen
            };

            TextBox textBox = new TextBox() { Left = 50, Top = 20, Width = 400 };
            Button confirmation = new Button() { Text = "Ok", Left = 350, Width = 100, Top = 70, DialogResult = DialogResult.OK };
            confirmation.Click += (sender, e) => { prompt.Close(); };

            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.AcceptButton = confirmation;

            DialogResult result = prompt.ShowDialog();

            // If the user clicked OK, try to parse the result
            if (result == DialogResult.OK)
            {
                if (int.TryParse(textBox.Text, out int wordCountGoal))
                {
                    goalWordCount = wordCountGoal;
                    
                    if (!init)
                    {
                        ThisAddIn_DocumentChange();

                    }

                }
                else
                {
                    MessageBox.Show("Invalid input. Please enter a number.");
                    PromptForWordCountGoal();
                }
            }
        }

        private void InitializeStatusBar()
        {
            statusBarForm = new Form();
            statusBarForm.Width = 600;
            statusBarForm.Height = 150;
            statusBarForm.FormBorderStyle = FormBorderStyle.SizableToolWindow;  
            statusBarForm.StartPosition = FormStartPosition.Manual;    
            statusBarForm.TopMost = true;                              
            statusBarForm.ShowInTaskbar = false;                       

            label = new Label();
            label.AutoSize = false;
            label.Width = 400;
            label.TextAlign = ContentAlignment.MiddleCenter;
            statusBarForm.Controls.Add(label);


            progressBar = new ProgressBar()
            {
                Width = 400,
                Height = 30,
                Maximum = goalWordCount,
                Style = ProgressBarStyle.Continuous,
                ForeColor = Color.Green,
                Top = label.Height,
            };
            changeCountButton = new Button() { Text = "Change Word Count", Left = 100, Width = 200, Top = 70, DialogResult = DialogResult.OK };
            changeCountButton.Click += (sender, e) => { PromptForWordCountGoal(); };

            //padding alignment
            int padding = 5; 
            progressBar.Left = (statusBarForm.Width - progressBar.Width) / 2;
            changeCountButton.Left = (statusBarForm.Width - changeCountButton.Width) / 2;
            label.Left = (statusBarForm.Width - label.Width) / 2;


            statusBarForm.Controls.Add(changeCountButton);
            statusBarForm.Controls.Add(progressBar);

            UpdateStatusBar(0, 0);
            //SetStatusBarPosition();
            statusBarForm.Show();
        }
        private void Application_WindowActivate(Word.Document Doc, Word.Window Wn)
        {
            // Update status bar position when Word window is activated or resized
            UpdateStatusBar(0, 0);
            SetStatusBarPosition();
        }

        private void SetStatusBarPosition()
        {
            // Get the Word window's position and dimensions
            int wordLeft = this.Application.Left;
            int wordTop = this.Application.Top;
            int wordWidth = this.Application.Width;
            int wordHeight = this.Application.Height;

            // Position the status bar below the Word window
            statusBarForm.Left = wordLeft;
            statusBarForm.Top = wordTop + wordHeight;
            statusBarForm.Width = wordWidth;
        }



        private void UpdateStatusBar(int count, double percentage)
        {
            if (label.InvokeRequired || progressBar.InvokeRequired)
            {
                label.Invoke(new Action(() => label.Text = $"Word Count Tracker: {count} words ({percentage:F1}%)"));
                progressBar.Invoke(new Action(() => progressBar.Value = count));
            }
            else
            {
                label.Text = $"Word Count Tracker: {count} words ({percentage:F1}%)";
                progressBar.Value = count;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Unsubscribe from events and close the form when shutting down
            this.Application.DocumentBeforeSave -= ThisAddIn_DocumentBeforeSave;
            this.Application.WindowActivate -= this.Application_WindowActivate;
            statusBarForm.Close();
        }

        #region VSTO generated code
        // Leave this VSTO-generated code as is
        // ... (code from ThisAddIn.Designer.cs)
        #endregion
    }
}