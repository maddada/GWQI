using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using CsvHelper;

namespace GWQI
{
    public partial class MainForm : Form
    {
        private string folderName = "";
        public MainForm()
        {
            InitializeComponent();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
        }

        public double calciumWeight = 0.03;
        public double chlorideWeight = 0.26;
        public double fluorideWeight = 0.26;
        public double magnesiumWeight = 0.03;
        public double sulphateWeight = 0.26;
        public double bicarbonateWeight = 0.06;
        public double sodiumWeight = 0.1;

        private void MainForm_Load(object sender, EventArgs e)
        {
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
        }

        private void button_About_Click(object sender, EventArgs e)
        {
            Form2_About form2About = new Form2_About();
            form2About.ShowDialog();
        }

        private void button_Help_Click(object sender, EventArgs e)
        {
            Form3_Help form3Help = new Form3_Help();
            form3Help.ShowDialog();
        }

        private void textBox_Calcium_TextChanged(object sender, EventArgs e)
        {
            recalculateGWQI();
        }

        private void textBox_Chloride_TextChanged(object sender, EventArgs e)
        {
            recalculateGWQI();
        }

        private void textBox_Fluoride_TextChanged(object sender, EventArgs e)
        {
            recalculateGWQI();
        }

        private void textBox_Magnesium_TextChanged(object sender, EventArgs e)
        {
            recalculateGWQI();
        }

        private void textBoxSulphate_TextChanged(object sender, EventArgs e)
        {
            recalculateGWQI();
        }

        private void textBox_Bicarbonate_TextChanged(object sender, EventArgs e)
        {
            recalculateGWQI();
        }

        private void textBox_Sodium_TextChanged(object sender, EventArgs e)
        {
            recalculateGWQI();
        }

        public void recalculateGWQI()
        {
            Double.TryParse(textBox_Calcium.Text, out double Calcium);
            Double.TryParse(textBox_Chloride.Text, out double Chloride);
            Double.TryParse(textBox_Fluoride.Text, out double Fluoride);
            Double.TryParse(textBox_Magnesium.Text, out double Magnesium);
            Double.TryParse(textBox_Sulphate.Text, out double Sulphate);
            Double.TryParse(textBox_Bicarbonate.Text, out double Bicarbonate);
            Double.TryParse(textBox_Sodium.Text, out double Sodium);

            //! Calcium Result
            double calciumRes = CalcCalcium(Calcium);
            double chlorideRes = CalcChloride(Chloride);
            double fluorideRes = CalcFluoride(Fluoride);
            double magnesiumRes = CalculateMagnesium(Magnesium);
            double sulphateRes = CalcSulphate(Sulphate);
            double bicarbonateRes = CalcBicarbonate(Bicarbonate);
            double sodiumRes = CalcSodium(Sodium);

            double gwqiResult = Math.Pow(calciumRes, calciumWeight) *
                         Math.Pow(chlorideRes, chlorideWeight) *
                         Math.Pow(fluorideRes, fluorideWeight) *
                         Math.Pow(magnesiumRes, magnesiumWeight) *
                         Math.Pow(sulphateRes, sulphateWeight) *
                         Math.Pow(bicarbonateRes, bicarbonateWeight) *
                         Math.Pow(sodiumRes, sodiumWeight);

            textBox_Result.Text = Math.Round(gwqiResult, 2).ToString();
        }

        private static double CalcCalcium(double Calcium)
        {
            // IF(AND(B2>0&&B2<=20),(100-((100-75)*((B2-0)/(20-0))));
            if (Calcium > 0 && Calcium <= 20)
                return (100 - ((100 - 75) * ((Calcium - 0) / (20 - 0))));
            // IF(AND(B2>20&&B2<=40),(75-((75-50)*((B2-20)/(40-20))));
            else if (Calcium > 20 && Calcium <= 40)
                return (75 - ((75 - 50) * ((Calcium - 20) / (40 - 20))));
            // IF(AND(B2>40&&B2<=60),(50-((50-25)*((B2-40)/(60-40))));
            else if (Calcium > 40 && Calcium <= 60)
                return (50 - ((50 - 25) * ((Calcium - 40) / (60 - 40))));
            // IF(AND(B2>60&&B2<=80),(25-((25-5)*((B2-60)/(80-60))));
            else if (Calcium > 60 && Calcium <= 80)
                return (25 - ((25 - 5) * ((Calcium - 60) / (80 - 60))));
            // IF(B2>80,5)))))
            else if (Calcium > 80)
                return 5;
            else
                return 0;
        }

        private static double CalcChloride(double Chloride)
        {
            // IF(AND(C2>0,C2<=62.5),(100-((100-75)*((C2-0)/(62.5-0)))),
            if (Chloride > 0 && Chloride <= 62.5)
                return (100 - ((100 - 75) * ((Chloride - 0) / (62.5 - 0))));
            // IF(AND(C2>62.5,C2<=125),(75-((75-50)*((C2-62.5)/(125-62.5)))),
            else if (Chloride > 62.5 && Chloride <= 125)
                return (75 - ((75 - 50) * ((Chloride - 62.5) / (125 - 62.5))));
            // IF(AND(C2>125,C2<=187.5),(50-((50-25)*((C2-125)/(187.5-125)))),
            else if (Chloride > 125 && Chloride <= 187.5)
                return (50 - ((50 - 25) * ((Chloride - 125) / (187.5 - 125))));
            // IF(AND(C2>187.5,C2<=250),(25-((25-5)*((C2-187.5)/(250-187.5)))),
            else if (Chloride > 187.5 && Chloride <= 250)
                return (25 - ((25 - 5) * ((Chloride - 187.5) / (250 - 187.5))));
            // IF(C2>250,5)))))
            else if (Chloride > 250)
                return 5;
            else
                return 0;
        }
        private static double CalculateMagnesium(double Magnesium)
        {
            //IF(AND(Magnesium>0&&Magnesium<=7.5),(100-((100-75)*((Magnesium-0)/(7.5-0))));
            if (Magnesium > 0 && Magnesium <= 7.5)
                return (100 - ((100 - 75) * ((Magnesium - 0) / (7.5 - 0))));
            //IF(AND(Magnesium>7.5&&Magnesium<=15),(75-((75-50)*((Magnesium-7.5)/(15-7.5))));
            else if (Magnesium > 7.5 && Magnesium <= 15)
                return (75 - ((75 - 50) * ((Magnesium - 7.5) / (15 - 7.5))));
            //IF(AND(Magnesium>15&&Magnesium<=22.5),(50-((50-25)*((Magnesium-15)/(22.5-15))));
            else if (Magnesium > 15 && Magnesium <= 22.5)
                return (50 - ((50 - 25) * ((Magnesium - 15) / (22.5 - 15))));
            //IF(AND(Magnesium>22.5&&Magnesium<=30),(25-((25-5)*((Magnesium-22.5)/(30-22.5))));
            else if (Magnesium > 22.5 && Magnesium <= 30)
                return (25 - ((25 - 5) * ((Magnesium - 22.5) / (30 - 22.5))));
            //IF(Magnesium>30,5)))))
            else if (Magnesium > 30)
                return 5;
            else
                return 0;
        }


        private double CalcFluoride(double Fluoride)
        {
            // IF(AND(D21>0 && D21<=0.375),(100-((100-75)*((D21-0)/(0.375-0))));
            if (Fluoride > 0 && Fluoride <= 0.375)
                return (100 - ((100 - 75) * ((Fluoride - 0) / (0.375 - 0))));
            // IF(AND(D21>0.375 && D21<=0.75),(75-((75-50)*((D21-0.375)/(0.75-0.375))));
            else if (Fluoride > 0.375 && Fluoride <= 0.75)
                return (75 - ((75 - 50) * ((Fluoride - 0.375) / (0.75 - 0.375))));
            // IF(AND(D21>0.75 && D21<=1.125),(50-((50-25)*((D21-0.75)/(1.125-0.75))));
            else if (Fluoride > 0.75 && Fluoride <= 1.125)
                return (50 - ((50 - 25) * ((Fluoride - 0.75) / (1.125 - 0.75))));
            // IF(AND(D21>1.125 && D21<=1.5),(25-((25-5)*((D21-1.125)/(1.5-1.125))));
            else if (Fluoride > 1.125 && Fluoride <= 1.5)
                return (25 - ((25 - 5) * ((Fluoride - 1.125) / (1.5 - 1.125))));
            // IF(D21>1.5,5)))))
            else if (Fluoride > 1.5)
                return 5;
            else
                return 0;
        }

        private double CalcSulphate(double Sulphate)
        {
            //IF(AND(Sulphate>0&&Sulphate<=62.5),(100-((100-75)*((Sulphate-0)/(62.5-0))))
            if (Sulphate > 0 && Sulphate <= 62.5)
                return (100 - ((100 - 75) * ((Sulphate - 0) / (62.5 - 0))));
            //IF(AND(Sulphate>62.5&&Sulphate<=125),(75-((75-50)*((Sulphate-62.5)/(125-62.5))))
            else if (Sulphate > 62.5 && Sulphate <= 125)
                return (75 - ((75 - 50) * ((Sulphate - 62.5) / (125 - 62.5))));
            //IF(AND(Sulphate>125&&Sulphate<=187.5),(50-((50-25)*((Sulphate-125)/(187.5-125))))
            else if (Sulphate > 125 && Sulphate <= 187.5)
                return (50 - ((50 - 25) * ((Sulphate - 125) / (187.5 - 125))));
            //IF(AND(Sulphate>187.5&&Sulphate<=250),(25-((25-5)*((Sulphate-187.5)/(250-187.5))))
            else if (Sulphate > 187.5 && Sulphate <= 250)
                return (25 - ((25 - 5) * ((Sulphate - 187.5) / (250 - 187.5))));
            //IF(Sulphate>250,5)))))
            else if (Sulphate > 250)
                return 5;
            else
                return 0;
        }

        private double CalcSodium(double Sodium)
        {
            //IF(AND(Sodium > 0&& Sodium <= 37.5), (100 - ((100 - 75) * ((Sodium - 0) / (37.5 - 0))));
            if (Sodium > 0 && Sodium <= 37.5)
                return (100 - ((100 - 75) * ((Sodium - 0) / (37.5 - 0))));
            //IF(AND(Sodium > 37.5 && Sodium <= 75), (75 - ((75 - 50) * ((Sodium - 37.5) / (75 - 37.5))));
            else if (Sodium > 37.5 && Sodium <= 75)
                return (75 - ((75 - 50) * ((Sodium - 37.5) / (75 - 37.5))));
            //IF(AND(Sodium > 75 && Sodium <= 112.5), (50 - ((50 - 25) * ((Sodium - 75) / (112.5 - 75)))); 
            else if (Sodium > 75 && Sodium <= 112.5)
                return (50 - ((50 - 25) * ((Sodium - 75) / (112.5 - 75))));
            //IF(AND(Sodium > 112.5 && Sodium <= 150), (25 - ((25 - 5) * ((Sodium - 112.5) / (150 - 112.5))));
            else if (Sodium > 112.5 && Sodium <= 150)
                return (25 - ((25 - 5) * ((Sodium - 112.5) / (150 - 112.5))));
            //IF(Sodium > 150, 5)))))
            else if (Sodium > 150)
                return 5;
            else
                return 0;
        }

        private double CalcBicarbonate(double Bicarbonate)
        {
            // IF(AND(Bicarbonate>0&&Bicarbonate<=75),(100-((100-75)*((Bicarbonate-0)/(75-0))));
            if (Bicarbonate > 0 && Bicarbonate <= 75)
                return (100 - ((100 - 75) * ((Bicarbonate - 0) / (75 - 0))));
            //IF(AND(Bicarbonate>75&&Bicarbonate<=150),(75-((75-50)*((Bicarbonate-75)/(150-75))));
            else if (Bicarbonate > 75 && Bicarbonate <= 150)
                return (75 - ((75 - 50) * ((Bicarbonate - 75) / (150 - 75))));
            //IF(AND(Bicarbonate>150&&Bicarbonate<=225),(50-((50-25)*((Bicarbonate-150)/(225-150))));
            else if (Bicarbonate > 150 && Bicarbonate <= 225)
                return (50 - ((50 - 25) * ((Bicarbonate - 150) / (225 - 150))));
            //IF(AND(Bicarbonate>225&&Bicarbonate<=300),(25-((25-5)*((Bicarbonate-225)/(300-225))));
            else if (Bicarbonate > 225 && Bicarbonate <= 300)
                return (25 - ((25 - 5) * ((Bicarbonate - 225) / (300 - 225))));
            //IF(Bicarbonate>300,5)))))
            else if (Bicarbonate > 300)
                return 5;
            else
                return 0;
        }

        private bool fileOpened = false;

        private void button_selectYourFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "CSV Files (*.csv)|*.csv";
            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = false;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string sFileName = openFileDialog.FileName;
                // string[] arrAllFiles = choofdlog.FileNames; //used when Multiselect = true     
                try
                {
                    using (var reader = new StreamReader(sFileName))
                    using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                    {
                        var records = csv.GetRecords<GWQIDataRow>().ToList();

                        for (int i = 0; i < records.Count(); i++)
                        {
                            var r = records[i];
                            Debug.WriteLine(r);

                            Double.TryParse(r.Calcium, out double Calcium);
                            Double.TryParse(r.Chloride, out double Chloride);
                            Double.TryParse(r.Fluoride, out double Fluoride);
                            Double.TryParse(r.Magnesium, out double Magnesium);
                            Double.TryParse(r.Sulphate, out double Sulphate);
                            Double.TryParse(r.Bicarbonate, out double Bicarbonate);
                            Double.TryParse(r.Sodium, out double Sodium);

                            //! Calcium Result
                            double calciumRes = CalcCalcium(Calcium);
                            double chlorideRes = CalcChloride(Chloride);
                            double fluorideRes = CalcFluoride(Fluoride);
                            double magnesiumRes = CalculateMagnesium(Magnesium);
                            double sulphateRes = CalcSulphate(Sulphate);
                            double bicarbonateRes = CalcBicarbonate(Bicarbonate);
                            double sodiumRes = CalcSodium(Sodium);

                            double gwqiResult = Math.Pow(calciumRes, calciumWeight) *
                                                Math.Pow(chlorideRes, chlorideWeight) *
                                                Math.Pow(fluorideRes, fluorideWeight) *
                                                Math.Pow(magnesiumRes, magnesiumWeight) *
                                                Math.Pow(sulphateRes, sulphateWeight) *
                                                Math.Pow(bicarbonateRes, bicarbonateWeight) *
                                                Math.Pow(sodiumRes, sodiumWeight);

                            r._ = " ";
                            r.GWQI = Math.Round(gwqiResult, 2).ToString();

                            if (!String.IsNullOrEmpty(r.GWQI) && r.GWQI.Trim() == "0")
                            {
                                r.Sodium = " ";
                                r.GWQI = " ";
                                r.Fluoride = " ";
                                r.Sulphate = " ";
                                r.Bicarbonate = " ";
                                r.Calcium = " ";
                                r.Chloride = " ";
                                r._ = " ";
                                r.Date = " ";
                                r.Magnesium = " ";
                            }
                        }



                        //? Save results after finishing

                        // Show the FolderBrowserDialog (to select a folder to save in)
                        DialogResult result = folderBrowserDialog1.ShowDialog();

                        try
                        {
                            if (result == DialogResult.OK)
                            {
                                folderName = folderBrowserDialog1.SelectedPath;

                                using (var writer = new StreamWriter(folderName + "\\GWQI Result.csv"))
                                {
                                    using (var csvWriter = new CsvWriter(writer, CultureInfo.InvariantCulture))
                                    {
                                        csvWriter.WriteRecords(records);
                                        string message =
                                            $"Result File Successfully Saved at:\n{folderName}\\GWQI Result.csv";
                                        string title = "Success";
                                        MessageBox.Show(message, title);

                                        string fileName = folderName + "\\GWQI Result.csv";
                                        var process = new System.Diagnostics.Process();
                                        process.StartInfo = new System.Diagnostics.ProcessStartInfo() { UseShellExecute = true, FileName = fileName };

                                        process.Start();
                                    }
                                }

                                //if (!fileOpened)
                                //{
                                //    // No file is opened, bring up openFileDialog in selected path.
                                //    openFileDialog1.InitialDirectory = folderName;
                                //    openFileDialog1.FileName = null;
                                //    openMenuItem.PerformClick();
                                //}
                            }
                        }
                        catch (Exception ex)
                        {
                            string message =
                                $"Failed to save Result File at:\n{folderName}\\GWQI Result.csv\n\nFile could be in use, close it and try again.";
                            string title = "Error";
                            MessageBox.Show(message, title);
                        }
                    }
                }
                catch (Exception ex)
                {
                    string message =
                        $"Error Opening File:\n\nFile could be in use, close it and try again.";
                    string title = "Error";
                    MessageBox.Show(message, title);
                }
            }



        }

        private FolderBrowserDialog folderBrowserDialog1;

        private void button_downloadStartingTemplate_Click(object sender, EventArgs e)
        {
            //string templateCsvStr = @"Date,Calcium,Chloride,Fluoride,Magnesium,Sulphate,Bicarbonate,Sodium
            //,0,0,0,0,0,0,0";

            var records = new List<GWQIDataRow>
            {
                new GWQIDataRow()
                {
                    Date = "2006-04-01",
                    Calcium = "0",
                    Chloride = "0",
                    Fluoride = "0",
                    Magnesium = "0",
                    Sulphate = "0",
                    Bicarbonate = "0",
                    Sodium = "0",
                    _ = " ",
                    GWQI = " "
                },
            };

            // Show the FolderBrowserDialog.
            DialogResult result = folderBrowserDialog1.ShowDialog();

            try
            {
                if (result == DialogResult.OK)
                {
                    folderName = folderBrowserDialog1.SelectedPath;

                    using (var writer = new StreamWriter(folderName + "\\GWQI Template.csv"))
                    {
                        using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                        {
                            csv.WriteRecords(records);
                            string message = $"Template File Successfully Saved at:\n{folderName}\\GWQI Template.csv";
                            string title = "Success";
                            MessageBox.Show(message, title);
                        }
                    }

                    //if (!fileOpened)
                    //{
                    //    // No file is opened, bring up openFileDialog in selected path.
                    //    openFileDialog1.InitialDirectory = folderName;
                    //    openFileDialog1.FileName = null;
                    //    openMenuItem.PerformClick();
                    //}
                }
            }
            catch (Exception ex)
            {
                string message = $"Failed to save Template File at:\n{folderName}\\GWQI Template.csv\n\nFile could be in use, close it and try again.";
                string title = "Error";
                MessageBox.Show(message, title);
            }

        }


    }

    class GWQIDataRow
    {
        public string Date { get; set; }
        public string Calcium { get; set; }
        public string Chloride { get; set; }
        public string Fluoride { get; set; }
        public string Magnesium { get; set; }
        public string Sulphate { get; set; }
        public string Bicarbonate { get; set; }
        public string Sodium { get; set; }
        public string _ { get; set; }
        public string GWQI { get; set; }
    }
}


