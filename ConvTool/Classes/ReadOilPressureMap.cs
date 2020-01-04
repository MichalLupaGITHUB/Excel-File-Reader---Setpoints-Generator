using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;
using System.Windows.Controls;

namespace ConvTool.Classes
{
    public class ReadOilPressureMap
    {
        #region Private Members

        private FileStream ReadOilPressureMapExcelFile;

        private string ReadOilPressureMap_ExcelFilePath;
        private string OilPressureMap;

        private  String[,] OilPressureMapTable;

        private StringBuilder OutOilPressureMap;

        #endregion Private Members

        #region Public Members

        public StringBuilder[] OilPressureMapRecords;

        #endregion Public Members

        #region Constructors
        public ReadOilPressureMap(string ExcelFilePath)
        {
            ReadOilPressureMap_ExcelFilePath = ExcelFilePath;
        }
        #endregion Constructors

        #region Public Methods

        /// <summary>
        /// Get Oil Pressure Map Data
        /// </summary>
        /// <param name="ExcelFilePath"></param>
        public void GetOilPressureMapData()
        {
            LoadExcel ExcelFile = new LoadExcel();
            ReadOilPressureMapExcelFile = ExcelFile.OpenExcelFile(ReadOilPressureMap_ExcelFilePath);

            try
            {
                using (var Reader = ExcelDataReader.ExcelReaderFactory.CreateReader(ReadOilPressureMapExcelFile))
                {
                    var Result = Reader.AsDataSet();
                    var DataGeneralTab = Result.Tables[1];

                    OilPressureMapTable = new String[DataGeneralTab.Rows.Count, DataGeneralTab.Columns.Count];
                    OilPressureMap = CalculateOilPressureMap(OilPressureMapTable, DataGeneralTab, OilPressureMap);             
                    OutOilPressureMap = new StringBuilder(OilPressureMap.Substring(0, OilPressureMap.IndexOf("\n\n")));
                }
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("Reading Oil Pressure Table Error!", "ConvTool - Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                Environment.Exit(8);
            }

            ExcelFile.CloseExcelFile();
        }
 
        /// <summary>
        /// Saving Oil Pressure Map as CSV File
        /// </summary>
        /// <param name="DestinationPath"></param>
        public void SaveOilPressureMapFile(string DestinationPath)
        {
            using (System.IO.StreamWriter OilPressureMapFile = new System.IO.StreamWriter(DestinationPath + "\\" + "OilPressureMap.csv"))
            {
                OilPressureMapFile.Write(OutOilPressureMap.ToString());              
                OilPressureMapFile.Close();
            }
        }

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Calculating Oil Pressure Map
        /// </summary>
        /// <param name="OilPressureMapTable"></param>
        /// <param name="DataGeneralTab"></param>
        /// <param name="OilPressureMap"></param>
        /// <returns></returns>
        private string CalculateOilPressureMap(String[,] OilPressureMapTable, System.Data.DataTable DataGeneralTab, string OilPressureMap)
        {
            OilPressureMapRecords = new StringBuilder[this.OilPressureMapTable.GetLength(0)];

            string OilPressureMapTemp = null; // variable for displaying on the grid

            for (int iRow = 1; iRow < DataGeneralTab.Rows.Count; iRow++)
            {
                // instructions for displaying on the grid
                if (OilPressureMapTemp != String.Empty)
                {
                    if (iRow != 1)
                    {
                        OilPressureMapRecords[iRow - 1] = new StringBuilder(OilPressureMapTemp + "\n");
                        OilPressureMapTemp = null;
                    }
                    else
                    {
                        OilPressureMapRecords[iRow - 1] = new StringBuilder("Temp / Speed");
                    }
                }
                else
                {
                    OilPressureMapRecords[iRow - 1] = null;
                    break;
                }
                // ----

                for (int iCol = 1; iCol < DataGeneralTab.Columns.Count; iCol++)
                {
                    if (DataGeneralTab.Rows[iRow][iCol].ToString() != String.Empty)
                    {
                        this.OilPressureMapTable[iRow - 1, iCol - 1] = DataGeneralTab.Rows[iRow][iCol].ToString();
                        this.OilPressureMapTable[0, 0] = String.Empty;
                        this.OilPressureMap += this.OilPressureMapTable[iRow - 1, iCol - 1] + ";";
                        OilPressureMapTemp += this.OilPressureMapTable[iRow - 1, iCol - 1] + "\t" +"\t";
                    }
                    else
                    {
                        this.OilPressureMapTable[iRow - 1, iCol - 1] = null;
                        this.OilPressureMap += "\n";                       
                        break;
                    }
                } 
            }

            return this.OilPressureMap;
        }

        #endregion Private Methods
    }
}
