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
    public class ReadINCA
    {
        #region Private Members

        private FileStream ReadINCAExcelFile;
        private readonly string ReadINCA_ExcelFilePath;

        private INCATable RecordTable;

        public  List<string> TableINCAChannels;
        public  List<string> TableINCAValues;
        public List<INCATable> INCATableList = new List<INCATable>();

        public XmlDocument INCAXML = new XmlDocument();

        #endregion Private Members

        #region Public Members

        public string MainHeader;
        public string NameVariableHeader;
        public string NameValueHeader;
        public string TableName;

        public struct INCATable
        {
            public string TableName { get; internal set; }
            public List<string> TableINCAChannels { get; private set; }
            public List<string> TableINCAValues { get; private set; }

            public INCATable(string tableName, List<string> TableINCAChannels, List<string> TableINCAValues)
            {
                TableName = tableName;
                this.TableINCAChannels = new List<string>();
                this.TableINCAValues = new List<string>();
            }
        }
      
        #endregion Public Members

        #region Constructors
        public ReadINCA(string ExcelFilePath)
        {
            ReadINCA_ExcelFilePath = ExcelFilePath;
        }
        #endregion Constructors

        #region Public Methods

        /// <summary>
        /// Get Setpoints Data
        /// </summary>
        /// <param name="ExcelFilePath"></param>
        public void GetINCAData()
        {
            LoadExcel ExcelFile = new LoadExcel();
            ReadINCAExcelFile = ExcelFile.OpenExcelFile(ReadINCA_ExcelFilePath);

            try
            {
                using (var Reader = ExcelDataReader.ExcelReaderFactory.CreateReader(ReadINCAExcelFile))
                {
                    var Result = Reader.AsDataSet();

                    GenerateINCAList(Result, INCATableList, RecordTable, TableName, MainHeader, NameVariableHeader, NameValueHeader);
                }
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("Reading INCA Tables Error!", "ConvTool - Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                Environment.Exit(10);
            }

            ExcelFile.CloseExcelFile();
        }

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Generating Setpoints List
        /// </summary>
        /// <param name="Result"></param>
        /// <param name="Setpoints"></param>
        /// <param name="Record"></param>
        /// <returns></returns>
        private void GenerateINCAList(System.Data.DataSet Result, List<INCATable> INCATableList, INCATable RecordTable, string TableName, string MainHeader, string NameVariableHeader, string NameValueHeader)
        {
            for (int iTab = 3; iTab < Result.Tables.Count; iTab++) // Tables
            {
                var DataGeneralTab = Result.Tables[iTab];
                this.TableName = DataGeneralTab.TableName.ToString();

                TableINCAChannels = new List<string>();
                TableINCAValues = new List<string>();

                this.RecordTable = new INCATable(this.TableName, TableINCAChannels, TableINCAValues);             
                this.INCATableList.Add(this.RecordTable);
                
                this.MainHeader = DataGeneralTab.Rows[0][8].ToString();
                this.NameVariableHeader = DataGeneralTab.Rows[1][8].ToString();
                this.NameValueHeader = DataGeneralTab.Rows[1][9].ToString();

                for (int iRow = 2; iRow < DataGeneralTab.Rows.Count; iRow++) // Rows
                {
                    if (DataGeneralTab.Rows[iRow][8].ToString() == "List")
                    {                       
                        break;
                    }
                   
                    this.RecordTable.TableINCAChannels.Add(DataGeneralTab.Rows[iRow][8].ToString());
                    this.RecordTable.TableINCAValues.Add(DataGeneralTab.Rows[iRow][9].ToString());
                }
            }
        }

        /// <summary>
        /// Generating XML Node From COndCase Tab
        /// </summary>
        /// <returns></returns>
        public XmlDocument GenerateINCAXMLNodes()
        {
            try
            {
                XmlNode GeneralTableXML = INCAXML.CreateElement("Tables");
                INCAXML.AppendChild(GeneralTableXML);

                foreach (INCATable item in INCATableList)
                {                  
                    XmlNode TableXML = INCAXML.CreateElement(RemoveForbiddenXMLChars(item.TableName));
                    GeneralTableXML.AppendChild(TableXML);

                    XmlNode INCAHeaderXML = INCAXML.CreateElement(RemoveForbiddenXMLChars(MainHeader));
                    TableXML.AppendChild(INCAHeaderXML);

                    XmlNode INCAVariableXML = INCAXML.CreateElement(RemoveForbiddenXMLChars(NameVariableHeader));
                    INCAHeaderXML.AppendChild(INCAVariableXML);

                    for (int iINCAEl = 0; iINCAEl < item.TableINCAChannels.Count; iINCAEl++)
                    {
                        if (item.TableINCAChannels[iINCAEl] == String.Empty)
                        {
                            break;
                        }

                        XmlNode ChannelXML = INCAXML.CreateElement(RemoveForbiddenXMLChars(item.TableINCAChannels[iINCAEl]));
                        ChannelXML.InnerText = item.TableINCAValues[iINCAEl];
                        INCAVariableXML.AppendChild(ChannelXML);
                    }
                }                       
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("Creating INCA XML Nodes Error!", "ConvTool - Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                Environment.Exit(14);
            }

            return INCAXML;
        }

        /// <summary>
        /// Removing Selected Forbidden Chars in XML Node Name
        /// </summary>
        /// <param name="CurrentCell"></param>
        /// <returns></returns>
        private string RemoveForbiddenXMLChars(string CurrentCell)
        {
            CurrentCell = Regex.Replace(CurrentCell, @"\s", "");
            Regex TempRegex = new Regex("[^a-zA-Z0-9_ -]");
            CurrentCell = TempRegex.Replace(CurrentCell, "");
            CurrentCell = new string(CurrentCell.Where(ch => System.Xml.XmlConvert.IsXmlChar(ch)).ToArray());          
            return CurrentCell;
        }

        #endregion Private Methods
    }

}

