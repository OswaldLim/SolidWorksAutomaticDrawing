using SolidWorksBatchDXF;
using StepToDrawing;
using System;
using System.IO;
using System.Windows.Forms;

namespace StepToDrawing
{
    public partial class GUI : Form
    {
        public GUI()
        {
            InitializeComponent();
            InitializeCustomComponents();
        }

        private TextBox inputFileTextBox;
        private Button inputFileButton;
        private TextBox outputFolderTextBox;
        private Button outputFolderButton;
        private Button runButton;

        private void InitializeCustomComponents()
        {
            // Input File Label
            Label inputFileLabel = new Label
            {
                Text = "Input File:",
                Left = 20,
                Top = 0,
                Width = 100
            };
            this.Controls.Add(inputFileLabel);

            // Input File TextBox
            inputFileTextBox = new TextBox
            {
                Left = 20,
                Top = 20,
                Width = 250
            };
            this.Controls.Add(inputFileTextBox);

            // Input File Button
            inputFileButton = new Button
            {
                Text = "Browse File...",
                Left = 280,
                Top = 18
            };
            inputFileButton.Click += InputFileButton_Click;
            this.Controls.Add(inputFileButton);

            // Output Folder Label
            Label outputFolderLabel = new Label
            {
                Text = "Output Folder:",
                Left = 20,
                Top = 50,
                Width = 100
            };
            this.Controls.Add(outputFolderLabel);

            // Output Folder TextBox
            outputFolderTextBox = new TextBox
            {
                Left = 20,
                Top = 70,
                Width = 250
            };
            this.Controls.Add(outputFolderTextBox);

            // Output Folder Button
            outputFolderButton = new Button
            {
                Text = "Browse Folder...",
                Left = 280,
                Top = 68
            };
            outputFolderButton.Click += OutputFolderButton_Click;
            this.Controls.Add(outputFolderButton);

            // Run Button
            runButton = new Button
            {
                Text = "Run",
                Left = 20,
                Top = 110,
                Width = 100
            };
            runButton.Click += RunButton_Click;
            this.Controls.Add(runButton);
        }

        private void InputFileButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter =
                    "SolidWorks & STEP Files (*.sldasm;*.sldprt;*.step;*.stp)|*.sldasm;*.sldprt;*.step;*.stp|All files (*.*)|*.*";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    inputFileTextBox.Text = openFileDialog.FileName;
                }
            }
        }


        private void OutputFolderButton_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowser = new FolderBrowserDialog())
            {
                if (folderBrowser.ShowDialog() == DialogResult.OK)
                {
                    outputFolderTextBox.Text = folderBrowser.SelectedPath;
                }
            }
        }

        private void RunButton_Click(object sender, EventArgs e)
        {
            string inputFile = inputFileTextBox.Text;
            string outputFolder = outputFolderTextBox.Text;

            if (string.IsNullOrWhiteSpace(inputFile) || string.IsNullOrWhiteSpace(outputFolder))
            {
                MessageBox.Show("Please select both input file and output folder.");
                return;
            }

            string ext = Path.GetExtension(inputFile).ToLower();
            if (ext != ".sldasm" && ext != ".sldprt" && ext != ".step" && ext != ".stp")
            {
                MessageBox.Show("Please select a valid SolidWorks assembly (.sldasm), part (.sldprt), or STEP (.step/.stp) file.");
                return;
            }

            // Run the backend logic
            MessageBox.Show($"Running process...\nInput: {inputFile}\nOutput: {outputFolder}");
            Logic.RunBatchExport(inputFile, outputFolder);
            MessageBox.Show($"Finished Running! \nInput: {inputFile}\nOutput: {outputFolder}");

        }
    }
}
