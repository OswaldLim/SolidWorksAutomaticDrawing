using System;
using System.Windows.Forms;
using StepToDrawing;

namespace SolidWorksBatchDXF
{
    class Program
    {

        [STAThread] // Required for Windows Forms
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Start your GUI form
            Application.Run(new GUI());
        }
    }
}



