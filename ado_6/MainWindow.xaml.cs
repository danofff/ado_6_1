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
using OfficeOpenXml;
using System.IO;
using ado_6.MODEL;

namespace ado_6
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    //
    public partial class MainWindow : Window
    {
        public static DATA database = new DATA();
        public MainWindow()
        {
            InitializeComponent();
        }

        private void SubmitButton_Click(object sender, RoutedEventArgs e)
        {
            DateTime from = new DateTime();
            DateTime to = new DateTime();
            if (From.SelectedDate != null)
            {
                from = (DateTime)From.SelectedDate;
            }
            if (To.SelectedDate != null && To.SelectedDate > From.SelectedDate)
            {
                to = (DateTime)To.SelectedDate;
            }
            string path = @"C:\Users\Вершковд\Documents\Visual Studio 2015\Projects\ado_6\ado_6\Template\OverallRepairList.xlsx";
            FileInfo fi = new FileInfo(path);

            ExcelPackage package = new ExcelPackage();
            using (package=new ExcelPackage(fi, true))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                List<TrackEvaluation> list = database.TrackEvaluation.Where(w => w.dCreateDate >= from && w.dCreateDate <= to).ToList();
                int startRow = 3;
                foreach (var item in list)
                {
                    worksheet.Cells[startRow, 1].Value = item.newEquipment?.TablesModel?.strName??"";
                    worksheet.Cells[startRow, 2].Value = item.intEvalutionNumber;
                    worksheet.Cells[startRow, 3].Value = item.dCreateDate;
                    worksheet.Cells[startRow, 4].Value = item.strSalesOrderNumber;
                    worksheet.Cells[startRow, 5].Value = item.dConfirmDate;
                    worksheet.Cells[startRow, 6].Value = item.dDateOfSale;
                    worksheet.Cells[startRow, 7].Value = item.dEstimatedDateArrival;
                    worksheet.Cells[startRow, 8].Value = item.dArrivalDate;
                    worksheet.Cells[startRow, 10].Value = item.floatETTR;
                    worksheet.Cells[startRow, 11].Value = item.floatELTR;
                    worksheet.Cells[startRow, 14].Value = item.TableEvaluationSysStatus.strStatysName;
                    worksheet.Cells[startRow, 19].Value = item.TrackEvaluationPart.Select(s => s.floatUnitCostAvia).Sum();
                    startRow++;                                   
                }
                package.Save();
                FileInfo newFile = new FileInfo("newData.xlsx");
                package.SaveAs(newFile);
                if (list.Count == 0)
                    MessageBox.Show("No data", "No Data", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                else
                    MessageBox.Show("Data is downloaded");
              
            }

            

        }
    }
}
