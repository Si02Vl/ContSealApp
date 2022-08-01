using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace ContSealApp
{ 
    public partial class InputForm1 : Form
    {
        public InputForm1()
        {
            InitializeComponent();
            startButton.Click += StartButton_Click;
        }
        public void StartButton_Click(object? sender, EventArgs e)
        {
            outputBox.Clear();
            totalContainersBox.Clear();
            testBox1.Clear();

            if (radioButtonGetFromFile.Checked)
            {
                ResultList(SortedListFromClient(), GetInfoFromExcel());
                totalContainersBox.Text += outputBox.Lines.Length - 1;
            }

            else if (radioButtonGetFromSite.Checked)
            {
                //метод получения данных с сайта  
                totalContainersBox.Text += outputBox.Lines.Length - 1;
            }   
        }

        public void WriteToExcel_Click(object sender, EventArgs e)
        {
            testBox1.Clear();
            WriteToExcel();
        }
        
        public List<Container> SortedListFromClient() //Result is sorted list of containers(containerFromClientList) got from client
        {
            var inputTextFromClient = Regex.Replace(inputBox.Text, @"\.", ",").Trim();
            var inputList = inputTextFromClient.Split('\n');
            var inputContainersList = new string[inputList.Length];
            var inputWeightsList = new double[inputList.Length];

            int weightMultiplier = int.Parse(weightMultiplierValueBox.Text);

            List<Container> containersFromClientList = new();

            for (int n = 0; n < inputList.Length; n++)
            {
                string[] temp1 = inputList[n].Split(new char[] { ' ', '\t' });
                inputContainersList[n] = temp1[0];
                inputWeightsList[n] = Double.Parse(temp1[1]) * weightMultiplier;

                Container containerFromClient = new(n, inputContainersList[n], "None", inputWeightsList[n]);
                containersFromClientList.Add(containerFromClient);
            }

            var containersFromClientOutputList = from p in containersFromClientList
                        orderby p.ContainerNumber
                        select p;

            return containersFromClientList;
        } 
        public List<Container> GetInfoFromExcel() //Result is sorted list of containers(containersFromFileList) got from Excel file
        {
            List<Container> containersFromFileList = new();

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Excel.Application objExcel = new();
                var objWorkBook = objExcel.Workbooks.Open(openFileDialog1.FileName);
                var objWorkSheet = (Excel.Worksheet)objWorkBook.Sheets[1];
                var containersRange = objWorkSheet.UsedRange.Columns["A"];
                var sealsRange = objWorkSheet.UsedRange.Columns["B"];
                
                var containersFromFileArray = (Array)containersRange.Value;
                if (containersFromFileArray == null)
                    return containersFromFileList;
                
                var containersFromFile = containersFromFileArray.OfType<object>().Select(o => o.ToString()).ToArray();
                
                var sealsFromFileArray = (Array)sealsRange.Value;
                if (sealsFromFileArray == null)
                    return containersFromFileList;
                
                var sealsFromFile = sealsFromFileArray.OfType<object>().Select(o => o.ToString()).ToArray();

                if (containersFromFileArray.Length != sealsFromFile.Length)
                    return containersFromFileList;

                for (int i = 0; i < containersFromFileArray.Length; i++)
                {
                    testBox1.Text += $@"{containersFromFile[i]}
";

                    testBox2.Text += $"{sealsFromFile[i]}\n";

                    Container containerFromFile = new(i, containersFromFile[i], sealsFromFile[i], 0.0);
                    containersFromFileList.Add(containerFromFile);
                    
                }
                Application.DoEvents();
                objExcel.Quit();
            }
            
            return containersFromFileList;
        } 
        public List<Container> ResultList(List<Container> containerFromClientList, List<Container> containersFromFileList)
        {
            List<Container> resultList = new();
            
            //Comparing two lists by container number and getting new list
            var result = (from c in containerFromClientList
                join f in containersFromFileList on c.ContainerNumber equals f.ContainerNumber
                select new
                {
                    c.Id,
                    f.ContainerNumber,
                    f.ContainerSeal,
                    c.ContainerWeight
                });

            foreach (var c in result)
            {
                Container newContainer = new(c.Id, c.ContainerNumber, c.ContainerSeal, c.ContainerWeight);
                resultList.Add(newContainer);
            }
            
            //Output in TextBox
             for (int i = 0; i < result.Count(); i++)
             {
                 outputBox.Text += $"{resultList[i].ContainerNumber} - {resultList[i].ContainerSeal} - {resultList[i].ContainerWeight}\r\n";
             }
            
            return resultList;
        }
        public void WriteToExcel()
        {
            try
            {
                //var arrayRange = ResultList(SortedListFromClient(), GetInfoFromExcel()).Count;
                var arrayRange = ResultList(SortedListFromClient(), GetInfoFromExcel()).Count;
                
                Excel.Application excelApp = new()
                {
                    Visible = true
                };

                // Create new book and sheet
                var workBook = excelApp.Workbooks.Add();
                var workSheet = (Excel.Worksheet)workBook.Sheets.Add();
            
            
                workSheet.Cells[1, 1] = @"Номер контейнера";
                workSheet.Cells[1, 2] = @"Вес";
                workSheet.Cells[1, 3] = @"Пломба";

                var headerRange = workSheet.Range["A1", "C1"];
                headerRange.Font.Bold = true;
                headerRange.Font.Color = ColorTranslator.ToOle(Color.Black);
                headerRange.Interior.Color = ColorTranslator.ToOle(Color.LightGreen);

                int i = 2; //Start from 2nd line    
                foreach (var j in ResultList(SortedListFromClient(), GetInfoFromExcel()))
                {
                    workSheet.Cells[i, 1] = j.ContainerNumber;
                    workSheet.Cells[i, 2] = j.ContainerWeight;
                    workSheet.Cells[i++, 3] = j.ContainerSeal;
                }
            
                workSheet.Columns.AutoFit();
            
                excelApp.Quit();
                MessageBox.Show(@"Done!");
            }
            catch (Exception exception)
            {
                MessageBox.Show(@"Ошибка - " + exception.Message);
            }
            
        }
    }
    public class Container
    {
        public int Id { get; set; }
        public string ContainerNumber { get; set; }
        public string ContainerSeal { get; set; }
        public double ContainerWeight { get; set; }

        public Container(int id, string containerNumber, string containerSeal, double containerWeight)
        {
            Id = id;
            ContainerNumber = containerNumber;
            ContainerSeal = containerSeal;
            ContainerWeight = containerWeight;
        }
    }
}
