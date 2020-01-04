using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using ExcelDataReader; // included EXT Library concerning reading from Excel File
using ConvTool.Classes;
using System.IO;
using System.Reflection;


namespace ConvTool
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        /*
        Error Table
        0 - OK
        1 - Too small number of Args!
        2 - Excel File does not exist!
        3 - Destination Directiry does not exist!
        4 - Internal program Error!
        5 - Loading Excel File Error! Excel File must be closed!
        6 - Reading CondCase Table Error!
        7 - Reading General Table Error!
        8 - Reading Oil Pressure Table Error!
        9 - Reading Setpoints Tables Error!
        10 - Reading INCA Tables Error!
        11 - Saving XML File Error!
        12 - Creating General XML Nodes Error!
        13 - Creating Cond Case XML Nodes Error!
        14 - Creating INCA XML Nodes Error!
        */
        public App()
        {
            AppDomain.CurrentDomain.AssemblyResolve += CurrrentDomain_AssemblyResolve;
        }

        /// <summary>
        /// Startup with Args - connected with attribute in xaml properties
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void App_Startup(object sender, StartupEventArgs e)
        {
            String FirstArg = String.Empty;
            String SecondArg = String.Empty;

            bool AppMode = false;

            //Checking nr of Args
            if (e.Args.Length == 2)
            {
                FirstArg = e.Args[0];
                SecondArg = e.Args[1];
            }
            else if (e.Args.Length == 0)
            {
                // Creating MAINWINDOW object
                MainWindow mainWindow = new MainWindow();

                mainWindow.InitializeComponent();
                mainWindow.Show();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Too small number of Args!", "ConvTool - Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                Environment.Exit(1);
            }

            if (FirstArg != String.Empty && SecondArg != String.Empty)
            {
                //Checking existing of Excel File
                if (!File.Exists(FirstArg))
                {
                    System.Windows.Forms.MessageBox.Show("Excel File does not exist!", "ConvTool - Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    Environment.Exit(2);
                }

                // Checking existing of Directory Location
                if (!Directory.Exists(SecondArg))
                {
                    System.Windows.Forms.MessageBox.Show("Destination Directiry does not exist!", "ConvTool - Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    Environment.Exit(3);
                }

                // Creating MAINWINDOW object
                MainWindow mainWindow = new MainWindow
                {
                    ExcelFilePathArg = FirstArg,
                    DestFilesLocationArg = SecondArg
                };

                AppMode = true;

                mainWindow.InitializeComponent();
                mainWindow.Application(AppMode);
            }                   
        }

        /// <summary>
        /// Get Not-Native Missing dll files
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        private Assembly CurrrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            var currentAssembly = Assembly.GetExecutingAssembly();
            var requiredDllName = $"{(new AssemblyName(args.Name).Name)}.dll";
            var resource = currentAssembly.GetManifestResourceNames().Where(s => s.EndsWith(requiredDllName)).FirstOrDefault();

            if (resource != null)
            {
                using (var stream = currentAssembly.GetManifestResourceStream(resource))
                {
                    if (stream == null)
                    {
                        return null;
                    }

                    var block = new byte[stream.Length];
                    stream.Read(block,0 ,block.Length);
                    return Assembly.Load(block);
                }
            }
            else
            {
                return null;
            }
        }
    }
}
