using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ExcelDataReader; // included EXT Library concerning reading from Excel File
using System.Xml;
using System.Xml.Linq;
using ConvTool.Classes;
using System.IO;
using System.Windows.Forms;

// VERSION 1.1 (26.06.2019)

namespace ConvTool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string ExcelFilePathArg;
        public string DestFilesLocationArg;

        private string ExcelFilePath;
        private string DestFilesLocation;
        private string XMLString;

        private ExcelPath obExcelPath = new ExcelPath();

        private OpenFileDialog e_fileBrowsing = new OpenFileDialog();
        private FolderBrowserDialog e_folderBrowsing = new FolderBrowserDialog();

        private bool AppMode;

        private XmlNode GeneralNode;
        private XmlNode CondCaseNode;
        private XmlNode INCANode;

        public MainWindow()
        {
                       
        }

        /// <summary>
        /// Main function to reading from Excel File and saving output files
        /// </summary>
        public void Application(bool ApplicationMode)
        {
            if ((ApplicationMode == true))
            {
                // Assigning Pathes from Args
                ExcelFilePath = ExcelFilePathArg;
                DestFilesLocation = DestFilesLocationArg;
            }

                // Datas of ExcelPath
                string ExcelPathExtension = System.IO.Path.GetExtension(ExcelFilePath);
                string ExcelPathFileName = System.IO.Path.GetFileName(ExcelFilePath);
                string ExcelPathFileNameNoExtension = System.IO.Path.GetFileNameWithoutExtension(ExcelFilePath);
                string ExcelPathRoot = System.IO.Path.GetPathRoot(ExcelFilePath);

                 // Datas of DestFilesLocation
                string DestFilesLocationExtension = System.IO.Path.GetExtension(DestFilesLocation);
                string DestFilesLocationFileName = System.IO.Path.GetFileName(DestFilesLocation);
                string DestFilesLocationFileNameNoExtension = System.IO.Path.GetFileNameWithoutExtension(DestFilesLocation);
                string DestFilesLocationRoot = System.IO.Path.GetPathRoot(DestFilesLocation);

            try
                {
                    // ---------------------------------------------------------
                    // ------------------ READ GENERAL DATA --------------------
                    // ---------------------------------------------------------

                    ReadGeneral obReadGeneral = new ReadGeneral(ExcelFilePath);
                    obReadGeneral.GetGeneralData();
                    GeneralNode = obReadGeneral.GenerateGeneralXMLNodes();

                    HeaderDocumentation.Text = obReadGeneral.DocumentationHeaders[0].ToUpper();
                    HeaderProject.Text = obReadGeneral.DocumentationHeaders[1];
                    HeaderPSP.Text = obReadGeneral.DocumentationHeaders[2];
                    HeaderActuator.Text = obReadGeneral.DocumentationHeaders[3];
                    HeaderUser.Text = obReadGeneral.DocumentationHeaders[4];
                    HeaderMeasurementNumber.Text = obReadGeneral.DocumentationHeaders[5];

                    ProjectContent.Text = obReadGeneral.DocumentationContents[1];
                    PSPContent.Text = obReadGeneral.DocumentationContents[2];
                    ActuatorContent.Text = obReadGeneral.DocumentationContents[3];
                    UserContent.Text = obReadGeneral.DocumentationContents[4];
                    MeasurementNumberContent.Text = obReadGeneral.DocumentationContents[5];

                    HeaderASAM3_INCA_Configuration.Text = obReadGeneral.ASAM3_INCA_ConfigurationHeaders[0].ToUpper();
                    HeaderHostNameIP.Text = obReadGeneral.ASAM3_INCA_ConfigurationHeaders[1];
                    HeaderProjectPath.Text = obReadGeneral.ASAM3_INCA_ConfigurationHeaders[2];
                    HeaderDataSet.Text = obReadGeneral.ASAM3_INCA_ConfigurationHeaders[3];

                    HostNameIPContent.Text = obReadGeneral.ASAM3_INCA_ConfigurationContents[1];
                    ProjectPathContent.Text = obReadGeneral.ASAM3_INCA_ConfigurationContents[2];
                    DataSetContent.Text = obReadGeneral.ASAM3_INCA_ConfigurationContents[3];

                    HeaderBatteryVoltageControl.Text = obReadGeneral.BatteryVoltageControlHeaders[0].ToUpper();
                    HeaderBatteryVoltageControlEnabler.Text = obReadGeneral.BatteryVoltageControlHeaders[1];
                    HeaderINCAVoltageVariable.Text = obReadGeneral.BatteryVoltageControlHeaders[2];
                    HeaderConditioningVoltage.Text = obReadGeneral.BatteryVoltageControlHeaders[3];
                    HeaderAngleCodeOpening.Text = obReadGeneral.BatteryVoltageControlHeaders[4];
                    HeaderAngleCodeClosing.Text = obReadGeneral.BatteryVoltageControlHeaders[5];

                    BatteryVoltageControlEnablerContent.Text = obReadGeneral.BatteryVoltageControlContents[1];
                    INCAVoltageVariableContent.Text = obReadGeneral.BatteryVoltageControlContents[2];
                    ConditioningVoltageContent.Text = obReadGeneral.BatteryVoltageControlContents[3];
                    AngleCodeOpeningContent.Text = obReadGeneral.BatteryVoltageControlContents[4];
                    AngleCodeClosingContent.Text = obReadGeneral.BatteryVoltageControlContents[5];
                    // ---------------------------------------------------------

                    // ---------------------------------------------------------
                    // ------------------ READ OIL PRESSURE MAP ----------------
                    // ---------------------------------------------------------

                    ReadOilPressureMap obReadOilPressureMap = new ReadOilPressureMap(ExcelFilePath);
                    obReadOilPressureMap.GetOilPressureMapData();
                    obReadOilPressureMap.SaveOilPressureMapFile(DestFilesLocation);

                    // Displaying Cond Case Data on the grid
                    OilPressureMapView.ItemsSource = obReadOilPressureMap.OilPressureMapRecords;
                    // ---------------------------------------------------------

                    // ---------------------------------------------------------
                    // ------------------ READ COND CASE DATA ------------------
                    // ---------------------------------------------------------

                    ReadCondCase obReadCondCase = new ReadCondCase(ExcelFilePath);
                    obReadCondCase.GetCondCaseData();
                    CondCaseNode = obReadCondCase.GenerateCondCaseXMLNodes();

                    // Displaying Cond Case Data on the grid
                    HeaderSpeed1.Text = obReadCondCase.HeaderSpeed1;
                    HeaderSpeed2.Text = obReadCondCase.HeaderSpeed2;

                    Speed1Content.Text = obReadCondCase.Speed1Content;
                    Speed2Content.Text = obReadCondCase.Speed2Content;

                    HeaderTemp1.Text = obReadCondCase.HeaderTemp1;
                    HeaderTemp2.Text = obReadCondCase.HeaderTemp2;

                    Temp1Content.Text = obReadCondCase.Temp1Content;
                    Temp2Content.Text = obReadCondCase.Temp2Content;
                    // ---------------------------------------------------------

                    // ---------------------------------------------------------
                    // ---------------------- READ INCA ------------------------
                    // ---------------------------------------------------------

                    ReadINCA obReadINCA = new ReadINCA(ExcelFilePath);
                    obReadINCA.GetINCAData();
                    INCANode = obReadINCA.GenerateINCAXMLNodes();
                    // ---------------------------------------------------------

                    // ---------------------------------------------------------
                    // ------------------ READ SETPOINTS -----------------------
                    // ---------------------------------------------------------

                    ReadSetpoints obReadSetpoints = new ReadSetpoints(ExcelFilePath);
                    obReadSetpoints.GetSetpointsData();
                    obReadSetpoints.SaveSetpointsFile(DestFilesLocation);

                    // Displaying Setpoints Data on the grid
                    SetpointsView.ItemsSource = obReadSetpoints.OutSetpoints;
                    // ---------------------------------------------------------

                    // ---------------------------------------------------------
                    // ------------------ SAVE XML FILE ------------------------
                    // ---------------------------------------------------------

                    SaveXMLFile obSaveXMLFile = new SaveXMLFile();
                    obSaveXMLFile.CreateXMLFile(GeneralNode, CondCaseNode, INCANode);
                    obSaveXMLFile.SaveXMLNodesFile(DestFilesLocation);
                    XMLString = obSaveXMLFile.ConvertXMLDocumentToString();

                    // Displaying XML Tree on the grid               
                    XMLView.Text = XMLString;
                    // ---------------------------------------------------------

                    if (ApplicationMode == true)
                    {
                        // Returned Value 0 when program has finished correct
                        Environment.Exit(0);
                    }
                }
                catch (Exception)
                {
                    System.Windows.Forms.MessageBox.Show("Internal Program Error!", "ConvTool - Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    Environment.Exit(4);
                }
            }

        private void ExcelFilePath_Click(object sender, RoutedEventArgs e)
        {          
            if (e_fileBrowsing.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ExcelFile_Path.Text = e_fileBrowsing.FileName;
                ExcelFilePath = obExcelPath.GetExcelPath(ExcelFile_Path.Text);
            }

            else
            {
                System.Windows.Forms.MessageBox.Show("Wrong Excel File Path!", "ConvTool - Warning", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }

        private void TargetLocationPath_Click(object sender, RoutedEventArgs e)
        {           
            if (e_folderBrowsing.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                TargetLocation_Path.Text = e_folderBrowsing.SelectedPath;
                DestFilesLocation = TargetLocation_Path.Text.ToString();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Wrong Target Location Path!", "ConvTool - Warning", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }

        private void Generate_Click(object sender, RoutedEventArgs e)
        {
            AppMode = false;

            if (ExcelFilePath != null && DestFilesLocation != null)
            {
                Application(AppMode);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Select Paths!", "ConvTool - Warning", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }                
    }
}
