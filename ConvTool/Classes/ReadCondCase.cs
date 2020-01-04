using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ExcelDataReader;
using System.Xml;
using System.Xml.Linq;
using System.Text.RegularExpressions;

namespace ConvTool.Classes
{
    public class ReadCondCase
    {
        #region Private Members

        private FileStream ReadCondCaseExcelFile;

        private string ReadCondCase_ExcelFilePath;

        private string TableName;

        private XmlDocument CondCaseXML = new XmlDocument();

        #endregion Private Members

        #region Public Members

        // Speed Headers
        public string HeaderSpeed1;
        public string HeaderSpeed2;

        // Speed Contents
        public string Speed1Content;
        public string Speed2Content;

        // Temperature Headers
        public string HeaderTemp1;
        public string HeaderTemp2;

        // Temperature Contents
        public string Temp1Content;
        public string Temp2Content;

        #endregion Public Members

        #region Constructors

        public ReadCondCase(string ExcelFilePath)
        {
            ReadCondCase_ExcelFilePath = ExcelFilePath;
        }
        #endregion Constructors

        #region Private Methods
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
        /// Get Conditioning Case Data
        /// </summary>
        public void GetCondCaseData()
        {
            LoadExcel ExcelFile = new LoadExcel();
            ReadCondCaseExcelFile = ExcelFile.OpenExcelFile(ReadCondCase_ExcelFilePath);

            try
            {
                using (var Reader = ExcelDataReader.ExcelReaderFactory.CreateReader(ReadCondCaseExcelFile))
                {
                    var Result = Reader.AsDataSet();
                    var DataGeneralTab = Result.Tables[2];

                    TableName = DataGeneralTab.TableName.ToString();

                    HeaderSpeed1 = DataGeneralTab.Rows[0][2].ToString();
                    HeaderSpeed2 = DataGeneralTab.Rows[0][8].ToString();

                    Speed1Content = DataGeneralTab.Rows[3][2].ToString();
                    Speed2Content = DataGeneralTab.Rows[3][8].ToString();

                    HeaderTemp1 = DataGeneralTab.Rows[14][0].ToString();
                    HeaderTemp2 = DataGeneralTab.Rows[25][0].ToString();

                    Temp1Content = DataGeneralTab.Rows[14][1].ToString();
                    Temp2Content = DataGeneralTab.Rows[25][1].ToString();
                }
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("Reading CondCase Table Error!", "ConvTool - Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                Environment.Exit(6);
            }

            ExcelFile.CloseExcelFile();
        }
        /// <summary>
        /// Generating XML Node From CondCase Tab
        /// </summary>
        /// <returns></returns>
        public XmlDocument GenerateCondCaseXMLNodes()
        {
            HeaderSpeed1 = RemoveForbiddenXMLChars(HeaderSpeed1);
            HeaderSpeed2 = RemoveForbiddenXMLChars(HeaderSpeed2);

            HeaderTemp1 = RemoveForbiddenXMLChars(HeaderTemp1);
            HeaderTemp2 = RemoveForbiddenXMLChars(HeaderTemp2);

            try
            {
                // Table Node
                XmlNode TableNameXML = CondCaseXML.CreateElement(RemoveForbiddenXMLChars(TableName));
                CondCaseXML.AppendChild(TableNameXML);

                XmlNode VariablesXML = CondCaseXML.CreateElement("Variables");
                TableNameXML.AppendChild(VariablesXML);

                // Nodes
                XmlNode HeaderSpeed1XML = CondCaseXML.CreateElement(HeaderSpeed1);
                HeaderSpeed1XML.InnerText = Speed1Content;
                VariablesXML.AppendChild(HeaderSpeed1XML);

                XmlNode HeaderSpeed2XML = CondCaseXML.CreateElement(HeaderSpeed2);
                HeaderSpeed2XML.InnerText = Speed2Content;
                VariablesXML.AppendChild(HeaderSpeed2XML);

                XmlNode HeaderTemp1XML = CondCaseXML.CreateElement(HeaderTemp1);
                HeaderTemp1XML.InnerText = Temp1Content;
                VariablesXML.AppendChild(HeaderTemp1XML);

                XmlNode HeaderTemp2XML = CondCaseXML.CreateElement(HeaderTemp2);
                HeaderTemp2XML.InnerText = Temp2Content;
                VariablesXML.AppendChild(HeaderTemp2XML);
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("Creating CondCase XML Nodes Error!", "ConvTool - Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                Environment.Exit(13);
            }

            return CondCaseXML;
        }

        #endregion Public Methods
    }
}

