using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;
using System.Xml;
using System.Xml.Linq;
using System.Text.RegularExpressions;

namespace ConvTool.Classes
{
    public class ReadGeneral
    {
        #region Private Members

        private FileStream ReadGeneralExcelFile;

        private string ReadGeneral_ExcelFilePath;
        private string TableName;

        private XmlDocument GeneralXML = new XmlDocument();

        #endregion Private Members

        #region Public Members

        // Documentation Headers
        public string[] DocumentationHeaders = new string[6];      
        /*
        Documentation
        Project
        PSP
        Actuator
        User
        MeasurementNumber
        */

        // Documentation Contents
        public string[] DocumentationContents = new string[6];
        /*
        Project
        PSP
        Actuator
        User
        MeasurementNumber
        */

        // ASAM3 / INCA configuration Headers
        public string[] ASAM3_INCA_ConfigurationHeaders = new string[4];
        /*
        ASAM3INCAConfiguration;
        HostNameIP;
        ProjectPath;
        DataSet;
        */

        // ASAM3 / INCA configuration Contents
        public string[] ASAM3_INCA_ConfigurationContents = new string[4];
        /*
        HostNameIP;
        ProjectPath;
        DataSet;
        */

        // Battery Voltahe Control Headers
        public string[] BatteryVoltageControlHeaders = new string[6];
        /*
        BatteryVoltageControl;
        BatteryVoltageControlEnabler;
        INCAVoltageVariable;
        ConditioningVoltage;
        AngleCodeOpening;
        AngleCodeClosing;
        */

        // Battery Voltage Control Contents
        public string[] BatteryVoltageControlContents = new string[6];
        /*
        BatterVoltageControlEnabler;
        INCAVoltageVariable;
        ConditioningVoltage;
        AngleCodeOpening;
        HeaderAngleCodeClosing;
        */

        #endregion Public Members

        #region Constructors
        public ReadGeneral(string ExcelFilePath)
        {
            ReadGeneral_ExcelFilePath = ExcelFilePath;
        }
        #endregion Constructors

        #region Private Methods
        /// <summary>
        /// Removing Selected Forbidden Chars in XML Node Name
        /// </summary>
        /// <param name="CurrentTableHeaders"></param>
        /// <returns></returns>
        private string[] RemoveForbiddenXMLChars(string [] CurrentTableHeaders)
        {
            for (int i = 0; i < CurrentTableHeaders.Length; i++)
            {
                CurrentTableHeaders[i] = Regex.Replace(CurrentTableHeaders[i], @"\s", "");
                Regex TempRegex = new Regex("[^a-zA-Z0-9 -]");
                CurrentTableHeaders[i] = TempRegex.Replace(CurrentTableHeaders[i], "");
                CurrentTableHeaders[i] = new string(CurrentTableHeaders[i].Where(ch => System.Xml.XmlConvert.IsXmlChar(ch)).ToArray());
            }

            return CurrentTableHeaders;
        }

        /// <summary>
        /// Removing Selected Forbidden Chars in XML Node Name
        /// </summary>
        /// <param name="CurrentCell"></param>
        /// <returns></returns>
        private string RemoveForbiddenXMLChars(string CurrentCell)
        {
            CurrentCell = Regex.Replace(CurrentCell, @"\s", "");
            Regex TempRegex = new Regex("[^a-zA-Z0-9 -]");
            CurrentCell = TempRegex.Replace(CurrentCell, "");
            CurrentCell = new string(CurrentCell.Where(ch => System.Xml.XmlConvert.IsXmlChar(ch)).ToArray());

            return CurrentCell;
        }

        #endregion Private Methods 

        #region Public Methods

        /// <summary>
        /// Get General Data
        /// </summary>
        public void GetGeneralData()
        {
            LoadExcel ExcelFile = new LoadExcel();
            ReadGeneralExcelFile = ExcelFile.OpenExcelFile(ReadGeneral_ExcelFilePath);

            try
            {
                using (var Reader = ExcelDataReader.ExcelReaderFactory.CreateReader(ReadGeneralExcelFile))
                {
                    var Result = Reader.AsDataSet();
                    var DataGeneralTab = Result.Tables[0];

                    TableName = DataGeneralTab.TableName.ToString();

                    for (int i = 1; i <= 6; i++)
                    {                       
                        DocumentationHeaders[i - 1] = DataGeneralTab.Rows[i][1].ToString();                      
                        DocumentationContents[i - 1] = DataGeneralTab.Rows[i][2].ToString();

                        BatteryVoltageControlHeaders[i - 1] = DataGeneralTab.Rows[i][7].ToString();
                        BatteryVoltageControlContents[i - 1] = DataGeneralTab.Rows[i][8].ToString();
                    }

                    for (int i = 8; i <= 11; i++)
                    {
                        ASAM3_INCA_ConfigurationHeaders[i - 8] = DataGeneralTab.Rows[i][1].ToString();
                        ASAM3_INCA_ConfigurationContents[i - 8] = DataGeneralTab.Rows[i][2].ToString();
                    }
                }
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("Reading General Table Error!", "ConvTool - Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                Environment.Exit(7);
            }

            ExcelFile.CloseExcelFile();
        }

        /// <summary>
        /// Generatring XML Node From General Tab
        /// </summary>
        /// <returns></returns>
        public XmlDocument GenerateGeneralXMLNodes()
        {
            DocumentationHeaders = RemoveForbiddenXMLChars(DocumentationHeaders);
            ASAM3_INCA_ConfigurationHeaders = RemoveForbiddenXMLChars(ASAM3_INCA_ConfigurationHeaders);
            BatteryVoltageControlHeaders = RemoveForbiddenXMLChars(BatteryVoltageControlHeaders);

            try
            {
                // Table Node
                XmlNode TableNameXML = GeneralXML.CreateElement(RemoveForbiddenXMLChars(TableName));
                GeneralXML.AppendChild(TableNameXML);

                // Documentation Node
                XmlNode DocumentationHeaderXML = GeneralXML.CreateElement(DocumentationHeaders[0]);
                TableNameXML.AppendChild(DocumentationHeaderXML);

                for (int iDocu = 1; iDocu < DocumentationHeaders.Length; iDocu++)
                {
                    XmlNode DocumentationNodeXML = GeneralXML.CreateElement(DocumentationHeaders[iDocu]);
                    DocumentationNodeXML.InnerText = DocumentationContents[iDocu];
                    DocumentationHeaderXML.AppendChild(DocumentationNodeXML);
                }

                // ASAM3_INCA_Configuration Node
                XmlNode ASAM3_INCA_ConfigurationHeaderXML = GeneralXML.CreateElement(ASAM3_INCA_ConfigurationHeaders[0]);
                TableNameXML.AppendChild(ASAM3_INCA_ConfigurationHeaderXML);

                for (int iASAM3 = 1; iASAM3 < ASAM3_INCA_ConfigurationHeaders.Length; iASAM3++)
                {
                    XmlNode ASAM3_INCA_ConfigurationNodeXML = GeneralXML.CreateElement(ASAM3_INCA_ConfigurationHeaders[iASAM3]);
                    ASAM3_INCA_ConfigurationNodeXML.InnerText = ASAM3_INCA_ConfigurationContents[iASAM3];
                    ASAM3_INCA_ConfigurationHeaderXML.AppendChild(ASAM3_INCA_ConfigurationNodeXML);
                }

                // Battery Voltage Control Node
                XmlNode BatteryVoltageControlHeaderXML = GeneralXML.CreateElement(BatteryVoltageControlHeaders[0]);
                TableNameXML.AppendChild(BatteryVoltageControlHeaderXML);

                for (int iBattery = 1; iBattery < BatteryVoltageControlHeaders.Length; iBattery++)
                {
                    XmlNode BatteryVoltageControlNodeXML = GeneralXML.CreateElement(BatteryVoltageControlHeaders[iBattery]);
                    BatteryVoltageControlNodeXML.InnerText = BatteryVoltageControlContents[iBattery];
                    BatteryVoltageControlHeaderXML.AppendChild(BatteryVoltageControlNodeXML);
                }
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("Creating General XML Nodes Error!", "ConvTool - Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                Environment.Exit(12);
            }

            return GeneralXML;
        }

        #endregion Public Methods
    }
}
