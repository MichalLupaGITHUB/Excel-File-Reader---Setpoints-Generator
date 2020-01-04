using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ExcelDataReader;

namespace ConvTool.Classes
{
    public class ReadSetpoints
    {
        #region Private Members

        private FileStream ReadSetpointsExcelFile;

        private int currTemperature;
        private int currSpeed;
        private int currOpeningAngle;
        private int currClosingAngles;
        private int NrOfMeasurementPoints = 0;
        private int currLastStep;

        private double currTOilCond;
        private double currTCoilCond;
        private double currTimeDuration;

        private string currTable;        
        private string Header;
        private string MorpheeVariables;
        private readonly string ReadSetpoints_ExcelFilePath;

        private bool TOilEqualTemp;
        private bool TOilEqualSpeed;
        private bool TCoilEqualTemp;
        private bool TCoilEqualSpeed;
        private bool RepeatFirstRowOP;
       
        private SetpointElements Record;

        private List<SetpointElements> Setpoints = new List<SetpointElements>();
        
        #endregion Private Members

        #region Public Members

        public StringBuilder[] OutSetpoints;
        public struct SetpointElements
        {
            public int Temperature { get; private set; }
            public int Speed { get; private set; }
            public int OpeningAngles { get; private set; }
            public int ClosingAngles { get; private set; }
            public double TOilCond { get; private set; }
            public double TCoilCond { get; private set; }
            public double TimeDuration { get; private set; }
            public int LastStep { get; internal set; }
            public string ExcelTab { get; private set; }

            public SetpointElements(int temperature, int speed, int openingAngles, int closingAngles, double toilCond, double tcoilCond, double timeDuration, int lastStep, string excelTab)
            {
                Temperature = temperature;
                Speed = speed;
                OpeningAngles = openingAngles;
                ClosingAngles = closingAngles;
                TOilCond = toilCond;
                TCoilCond = tcoilCond;
                TimeDuration = timeDuration;
                LastStep = lastStep;
                ExcelTab = excelTab;
            }
            
            public override string ToString()
            {
                return    this.Temperature
                        + this.Speed
                        + this.OpeningAngles
                        + this.ClosingAngles
                        + this.TOilCond
                        + this.TCoilCond
                        + this.TimeDuration
                        + this.LastStep
                        + this.ExcelTab;
            }                       
        }

        #endregion Public Members

        #region Constructors
        public ReadSetpoints(string ExcelFilePath)
        {
            ReadSetpoints_ExcelFilePath = ExcelFilePath;
        }
        #endregion Constructors

        #region Public Methods

        /// <summary>
        /// Get Setpoints Data
        /// </summary>
        /// <param name="ExcelFilePath"></param>
        public void GetSetpointsData()
        {
            LoadExcel ExcelFile = new LoadExcel();
            ReadSetpointsExcelFile = ExcelFile.OpenExcelFile(ReadSetpoints_ExcelFilePath);

            try
            {
                using (var Reader = ExcelDataReader.ExcelReaderFactory.CreateReader(ReadSetpointsExcelFile))
                {
                    var Result = Reader.AsDataSet();

                    Setpoints = GenerateSetpointsList(Result, Setpoints, Record);
                    Setpoints = Setpoints.OrderBy(Temp => Temp.Temperature).ToList();

                    OutSetpoints = new StringBuilder[Setpoints.Count];

                    for (int iOutSP = 0; iOutSP < Setpoints.Count; iOutSP++)
                    {                                              
                        OutSetpoints[iOutSP] = new StringBuilder((Setpoints[iOutSP].Temperature.ToString() + "\t" + Setpoints[iOutSP].Speed.ToString() + "\t" + Setpoints[iOutSP].OpeningAngles.ToString() + "\t" + Setpoints[iOutSP].ClosingAngles.ToString() + "\t" +Setpoints[iOutSP].TOilCond.ToString()+ "\t" + Setpoints[iOutSP].TCoilCond.ToString() + "\t" + Setpoints[iOutSP].TimeDuration.ToString() + "\t" + Setpoints[iOutSP].LastStep.ToString() + "\t" + Setpoints[iOutSP].ExcelTab).Replace(',', '.'));
                    }                                         
                }
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("Reading Setpoints Tables Error!", "ConvTool - Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                Environment.Exit(9);
            }

            ExcelFile.CloseExcelFile();
        }
        
        /// <summary>
        /// Save Setpoints File
        /// </summary>
        public void SaveSetpointsFile(string DestinationPath)
        {
            Header = "\";Setpoints[]\"";
            MorpheeVariables = "M_TD.Temperature" + "\t" + "M_TD.Motor_Speed" + "\t" + "M_TD.Opening_Angles" + "\t" + "M_TD.Closing_Angles" + "\t" + "M_TD.TOil_Cond" + "\t" + "M_TD.TCoil_Cond" + "\t" + "M_TD.Time_Duration" + "\t" + "M_TD.Last_Step" + "\t" + "M_TD.Assigned_Table";

            using (System.IO.StreamWriter SetpointsFile = new System.IO.StreamWriter(DestinationPath + "\\" + "Setpoints.txt"))
            {
                SetpointsFile.WriteLine(Header);
                SetpointsFile.WriteLine(MorpheeVariables);

                foreach (var line in OutSetpoints)
                {
                    SetpointsFile.WriteLine(line.ToString());
                }

                SetpointsFile.Close();
            }
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
        private List<SetpointElements> GenerateSetpointsList(System.Data.DataSet Result, List<SetpointElements> Setpoints, SetpointElements Record)
        {
            for (int iTab = 3; iTab < Result.Tables.Count; iTab++) // Tables
            {
                var DataGeneralTab = Result.Tables[iTab];

                for (int iTemp = 2; iTemp < DataGeneralTab.Rows.Count; iTemp++) // Temperature
                {
                    if (DataGeneralTab.Rows[iTemp][1].ToString() == String.Empty)
                    {
                        break;
                    }

                    for (int iSpeed = 2; iSpeed < DataGeneralTab.Rows.Count; iSpeed++) // Speed
                    {
                        RepeatFirstRowOP = true;

                        if (DataGeneralTab.Rows[iSpeed][3].ToString() == String.Empty)
                        {
                            break;
                        }

                        for (int iAngle = 2; iAngle < DataGeneralTab.Rows.Count; iAngle++) // OpeningAngles, Closing Angles
                        {
                            // Getting numbers of measurement points - quantity of records for specified temperature and speed
                            NrOfMeasurementPoints = CheckNrOfMeasurementPoints(NrOfMeasurementPoints, DataGeneralTab);

                            // Checking end of Operating Point and saving 1 in LastStep at the end of OP due to XPlayer requirements
                            if (DataGeneralTab.Rows[iAngle][5].ToString() == String.Empty || DataGeneralTab.Rows[iAngle][6].ToString() == String.Empty)
                            {
                                //this.Setpoints.RemoveAt(this.Setpoints.Count-1); // Removing last record due to saving 1 in LastStep

                                currLastStep = 1;

                                this.Record = new SetpointElements(currTemperature, currSpeed, 0, 0, currTOilCond, currTCoilCond, currTimeDuration, currLastStep, RemoveForbiddenXMLChars(currTable));
                                this.Setpoints.Add(this.Record);

                                NrOfMeasurementPoints = 0;

                                break;
                            }

                            // Conditions for reading TOil and TCoil Data
                            TOilEqualTemp = int.Parse(DataGeneralTab.Rows[iTemp][11].ToString()) == int.Parse(DataGeneralTab.Rows[iTemp][1].ToString());
                            TOilEqualSpeed = int.Parse(DataGeneralTab.Rows[1][iSpeed + 10].ToString()) == int.Parse(DataGeneralTab.Rows[iSpeed][3].ToString());
                            TCoilEqualTemp = int.Parse(DataGeneralTab.Rows[iTemp + 24][11].ToString()) == int.Parse(DataGeneralTab.Rows[iTemp][1].ToString());
                            TCoilEqualSpeed = int.Parse(DataGeneralTab.Rows[25][iSpeed + 10].ToString()) == int.Parse(DataGeneralTab.Rows[iSpeed][3].ToString());

                            if (TOilEqualTemp && TOilEqualSpeed && TCoilEqualTemp && TCoilEqualSpeed)
                            {
                                if (DataGeneralTab.Rows[iTemp][iSpeed + 10].ToString() == String.Empty)
                                {
                                    DataGeneralTab.Rows[iTemp][iSpeed + 10] = 0;
                                }

                                if (DataGeneralTab.Rows[iTemp + 24][iSpeed + 10].ToString() == String.Empty)
                                {
                                    DataGeneralTab.Rows[iTemp + 24][iSpeed + 10] = 0;
                                }

                                currTemperature = int.Parse(DataGeneralTab.Rows[iTemp][1].ToString());
                                currSpeed = int.Parse(DataGeneralTab.Rows[iSpeed][3].ToString());
                                currOpeningAngle = int.Parse(DataGeneralTab.Rows[iAngle][5].ToString());
                                currClosingAngles = int.Parse(DataGeneralTab.Rows[iAngle][6].ToString());
                                currTOilCond = Math.Round(double.Parse(DataGeneralTab.Rows[iTemp][iSpeed + 10].ToString()), 2);
                                currTCoilCond = Math.Round(double.Parse(DataGeneralTab.Rows[iTemp + 24][iSpeed + 10].ToString()), 2);
                                currTimeDuration = Math.Round(2 * 60 / (double.Parse(DataGeneralTab.Rows[iSpeed][3].ToString())) * 15, 2);
                                currLastStep = 0;
                                currTable = DataGeneralTab.TableName.ToString().Replace(" ", String.Empty);

                                this.Record = new SetpointElements(currTemperature, currSpeed, currOpeningAngle, currClosingAngles, currTOilCond, currTCoilCond, currTimeDuration, currLastStep, RemoveForbiddenXMLChars(currTable));
                                this.Setpoints.Add(this.Record);

                                if (RepeatFirstRowOP)
                                {
                                    this.Setpoints.Add(this.Record);
                                    RepeatFirstRowOP = false;
                                }

                                NrOfMeasurementPoints = 0;                               
                            }
                        }
                    }
                }
            }

            return this.Setpoints;
        }
        /// <summary>
        /// Checking Number Of Measurement Points
        /// </summary>
        /// <param name="NrOfMeasurementPoints"></param>
        /// <param name="RangeOpenAngles"></param>
        /// <param name="RangeCloseAngles"></param>
        /// <param name="DataGeneralTab"></param>
        /// <returns></returns>
        private int CheckNrOfMeasurementPoints(int NrOfMeasurementPoints, System.Data.DataTable DataGeneralTab)
        {
            for (int i = 2; i < DataGeneralTab.Rows.Count; i++) // 
            {
                if (DataGeneralTab.Rows[i][5].ToString() == String.Empty || DataGeneralTab.Rows[i][6].ToString() == String.Empty)
                {
                    break;
                }

                this.NrOfMeasurementPoints++;
            }
            return this.NrOfMeasurementPoints;
        }

        private string RemoveForbiddenXMLChars(string CurrentCell)
        {
            int indexChar = 0;
            Regex TempRegex = null;

            if (CurrentCell.Contains("("))
            {
                indexChar = CurrentCell.IndexOf("(");
                CurrentCell = Regex.Replace(CurrentCell, @"\s", "");
                TempRegex = new Regex("[^a-zA-Z0-9 -]");
                CurrentCell = TempRegex.Replace(CurrentCell, "");
                CurrentCell = CurrentCell.Insert(indexChar, "_");
            }
            else
            {
                CurrentCell = Regex.Replace(CurrentCell, @"\s", "");
                TempRegex = new Regex("[^a-zA-Z0-9 -]");
                CurrentCell = TempRegex.Replace(CurrentCell, "");             
            }
           
            return CurrentCell;
        }

        #endregion Private Methods
    }
}