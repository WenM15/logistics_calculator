using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using Microsoft.Win32;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

// (TODO) Might not need `bool abort` variable

namespace logistics_calculator
{
    using WeightMap = System.Collections.Generic.Dictionary<double, double?>;
    using HAWB_Map = System.Collections.Generic.Dictionary<int, System.Collections.Generic.Dictionary<double, double?>>;
    using QuesRegionMap = System.Collections.Generic.Dictionary<string, System.Collections.Generic.Dictionary<int, System.Collections.Generic.Dictionary<double, double?>>>;
    
    using WeightRangeMap = System.Collections.Generic.Dictionary<Weight, double>;
    using ServiceLevelMap = System.Collections.Generic.Dictionary<string, System.Collections.Generic.Dictionary<Weight, double>>;
    using RegionMap = System.Collections.Generic.Dictionary<string, System.Collections.Generic.Dictionary<string, System.Collections.Generic.Dictionary<Weight, double>>>;

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Dictionary<string, string> alias_map = new Dictionary<string, string>();

        private RegionMap ref_region_map = new RegionMap();
        private Dictionary<string, int> ref_header_map = new Dictionary<string, int>();
        private string target_service_level = "PRIORITY OVERNIGHT OR INTERNATIONAL PRIORITY";

        private QuesRegionMap ques_region_map = new QuesRegionMap();
        private Dictionary<string, int> ques_header_map = new Dictionary<string, int>();

        private string? outputFolderPath = null;

        public MainWindow()
        {
            InitializeComponent();            
            target_service_level = target_service_level.Replace(" ", "");
        }

        private void ProcessDictFile(object sender, RoutedEventArgs e)
        {
            string? filePath = GetFileFromDialog();
            if (filePath == null)
            {
                // sad path
                return;
            }

            foreach(string line in File.ReadLines(filePath))
            {
                string[] pair = line.Split('=');
                alias_map.Add(pair[0], pair[1]);
            }
        }

        private void ProcessRefFile(object sender, RoutedEventArgs e)
        {
            // Get file
            string? file_path = GetFileFromDialog();
            if (file_path == null)
            {
                // sad path
                return;
            }

            // Get reader
            FileStream stream = File.Open(file_path, FileMode.Open, FileAccess.Read);
            var reader = ExcelReaderFactory.CreateReader(stream);

            // Declare headers            
            ref_header_map["DestCountry/RegionCode"] = 0;
            ref_header_map["ServiceLevelShortDescription"] = 0;
            ref_header_map["PackageTypeCode"] = 0;
            ref_header_map["MinWeight"] = 0;
            ref_header_map["MaxWeight"] = 0;
            ref_header_map["Rate"] = 0;

            #region get column index based on header
            int found = 0;
            while (reader.Read())
            {
                for (int column = 0; column < reader.FieldCount; ++column)
                {
                    string? value = reader.GetValue(column)?.ToString();

                    if (value == null)
                    {
                        continue;
                    }

                    if (ref_header_map.ContainsKey(value))
                    {
                        ++found;
                        ref_header_map[value] = column;                        
                    }
                }
                if (found > 0)
                {
                    break;
                }
            }

            if (found != ref_header_map.Count)
            {
                return;
            }
            #endregion

            bool abort = false;
            while (reader.Read())
            {
                #region retrieve values
                string? service_level = reader.GetValue(ref_header_map["ServiceLevelShortDescription"])?.ToString();
                if (service_level == null)
                {   
                    // sad path
                    abort = true;
                    break;
                }
                service_level = service_level.Replace(" ", "");
                if (service_level != target_service_level)
                {
                    // unintended path
                    continue;
                }

                string? region = reader.GetValue(ref_header_map["DestCountry/RegionCode"])?.ToString();                                
                string? min_weight_str = reader.GetValue(ref_header_map["MinWeight"])?.ToString();
                string? max_weight_str = reader.GetValue(ref_header_map["MaxWeight"])?.ToString();
                string? rate_str = reader.GetValue(ref_header_map["Rate"])?.ToString();

                if (region == null || min_weight_str == null || max_weight_str == null || rate_str == null)
                {
                    // sad path
                    abort = true; 
                    break;
                }

                Weight weight = new Weight();
                if (double.TryParse(min_weight_str, out double min_weight) == false)
                {
                    // sad path
                    abort = true;
                    break;
                }
                if (double.TryParse(max_weight_str, out double max_weight) == false)
                {
                    // sad path
                    abort = true;
                    break;
                }
                weight.MinWeight = min_weight;
                weight.MaxWeight = max_weight;

                if (double.TryParse(rate_str, out double rate) == false)
                {
                    // sad path
                    abort = true;
                    break;
                }
                #endregion

                #region store values
                if (ref_region_map.ContainsKey(region) == false)
                {
                    ref_region_map.Add(region, new ServiceLevelMap());
                }
                if (ref_region_map[region].ContainsKey(service_level) == false)
                {
                    ref_region_map[region].Add(service_level, new WeightRangeMap());
                }
                if (ref_region_map[region][service_level].ContainsKey(weight) == false)
                {
                    ref_region_map[region][service_level].Add(weight, rate);
                }
                #endregion
            }

            if (abort == true)
            {
                // sad path
                return;
            }

            reader.Close();
            stream.Close();
        }

        private void ProcessQuesFile(object sender, RoutedEventArgs e)
        {
            // Get file
            string? file_path = GetFileFromDialog();
            if (file_path == null)
            {
                // sad path
                return;
            }

            // Get reader
            FileStream stream = File.Open(file_path, FileMode.Open, FileAccess.Read);
            var reader = ExcelReaderFactory.CreateReader(stream);

            // Declare headers
            ques_header_map["Ship-ToID"] = 0;
            ques_header_map["HAWB/BOL"] = 0;
            ques_header_map["GW"] = 0;

            #region get column index based on header
            int found = 0;
            while (reader.Read())
            {
                for (int column = 0; column < reader.FieldCount; ++column)
                {
                    string? value = reader.GetValue(column)?.ToString();

                    if (value == null)
                    {
                        continue;
                    }

                    if (ques_header_map.ContainsKey(value))
                    {
                        ++found;
                        ques_header_map[value] = column;
                    }
                }
                if (found > 0)
                {
                    break;
                }
            }

            if (found != ques_header_map.Count)
            {
                return;
            }
            #endregion

            bool abort = false;
            while (reader.Read())
            {
                #region retrieve values
                string? ship_to_id = reader.GetValue(ques_header_map["Ship-ToID"])?.ToString();
                string? hawb_str = reader.GetValue(ques_header_map["HAWB/BOL"])?.ToString();
                string? weight_str = reader.GetValue(ques_header_map["GW"])?.ToString();

                if (ship_to_id == null || hawb_str == null || weight_str == null)
                {
                    // sad path
                    abort = true;
                    break;
                }

                if (int.TryParse(hawb_str, out int hawb) == false)
                {
                    // sad path
                    abort = true;
                    break;
                }
                if (double.TryParse(weight_str, out double weight) == false)
                {
                    // sad path
                    abort = true;
                    break;
                }
                #endregion

                #region store values
                if (ques_region_map.ContainsKey(ship_to_id) == false)
                {
                    ques_region_map.Add(ship_to_id, new HAWB_Map());
                }
                if (ques_region_map[ship_to_id].ContainsKey(hawb) == false)
                {
                    ques_region_map[ship_to_id].Add(hawb, new WeightMap());
                }
                if (ques_region_map[ship_to_id][hawb].ContainsKey(weight) == false)
                {
                    ques_region_map[ship_to_id][hawb].Add(weight, null);
                }
                #endregion
            }
            if (abort == true)
            {
                return;
            }

            reader.Close();
            stream.Close();
        }

        private string? GetFileFromDialog()
        {
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == false)
            {
                return null;
            }
            return dialog.FileName;
        }

        private string? GetFolderFromDialog()
        {
            OpenFolderDialog dialog = new OpenFolderDialog();
            if (dialog.ShowDialog() == false)
            {
                return null;
            }
            return dialog.FolderName;
        }

        private void SelectOutputFolder(object sender, RoutedEventArgs e)
        {
            outputFolderPath = GetFolderFromDialog();
        }

        private void Start(object sender, RoutedEventArgs e)
        {         
            if (outputFolderPath == null || outputFile1TextBox.Text == null)
            {
                // sad path
                return;
            }

            StartPart1();
            StartPart2();

        }

        private void StartPart1()
        {
            #region ques1
            // Initialize excel file
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Sheet1");

            // Add header to sheet
            int row = 1;
            int col = 1;
            foreach (string? header in ques_header_map.Keys)
            {
                worksheet.Cell(row, col).Value = header;
                ++col;
            }
            worksheet.Cell(row, col).Value = "Cost";
            ++col;
            worksheet.Cell(row, col).Value = "Compounded cost";
            ++row;
            col = 1;

            // get individual cost
            foreach (var dest_hawb_pair in ques_region_map)
            {
                string ship_to_id = dest_hawb_pair.Key;
                HAWB_Map hawb_map = dest_hawb_pair.Value;

                string region;
                if (alias_map.ContainsKey(ship_to_id))
                {
                    region = alias_map[ship_to_id];
                }
                else
                {
                    // sad path
                    break;
                }

                foreach (var hawb_weight_pair in hawb_map)
                {
                    int hawb = hawb_weight_pair.Key;
                    WeightMap weight_map = hawb_weight_pair.Value;
                    double cost_sum = 0;

                    List<double> weights = weight_map.Keys.ToList();
                    for (int i = 0; i < weights.Count; ++i)
                    {
                        double weight = weights[i];

                        // Search weight range, continue to process once found
                        WeightRangeMap weight_range_map = ref_region_map[region][target_service_level];
                        foreach (KeyValuePair<Weight, double> weight_rate_pair in weight_range_map)
                        {
                            double min_weight = weight_rate_pair.Key.MinWeight;
                            double max_weight = weight_rate_pair.Key.MaxWeight;
                            double rate = weight_rate_pair.Value;

                            if ((weight >= min_weight) && (weight <= max_weight))
                            {
                                double cost = rate * weight;

                                weight_map[weight] = cost;
                                cost_sum += cost;

                                worksheet.Cell(row, col).Value = ship_to_id;
                                ++col;
                                worksheet.Cell(row, col).Value = hawb;
                                ++col;
                                worksheet.Cell(row, col).Value = weight;
                                ++col;
                                worksheet.Cell(row, col).Value = cost;
                                ++col;
                                worksheet.Cell(row, col).Value = cost_sum;

                                ++row;
                                col = 1;
                                break;
                            }
                        }
                    }
                }
            }

            // Write excel file to disk
            string fullPath = System.IO.Path.Combine(outputFolderPath, outputFile1TextBox.Text);
            workbook.SaveAs($"{fullPath}.xlsx");

            workbook.Dispose();
            #endregion
        }

        private void StartPart2()
        {
            // Initialize excel file
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Sheet1");

            // Add header to sheet
            int row = 1;
            int col = 1;

            worksheet.Cell(row, col).Value = "Ship-ToID";
            ++col;
            worksheet.Cell(row, col).Value = "HAWB/BOL";
            ++col;
            worksheet.Cell(row, col).Value = "Total weight";
            ++col;
            worksheet.Cell(row, col).Value = "Total cost (no consolidation)";
            ++col;
            worksheet.Cell(row, col).Value = "Total cost (consolidation)";
            ++row;
            col = 1;

            foreach (var dest_hawb_pair in ques_region_map)
            {
                string ship_to_id = dest_hawb_pair.Key;
                HAWB_Map hawb_map = dest_hawb_pair.Value;

                string region;
                if (alias_map.ContainsKey(ship_to_id))
                {
                    region = alias_map[ship_to_id];
                }
                else
                {
                    // sad path
                    break;
                }

                foreach (var hawb_weight_pair in hawb_map)
                {
                    int hawb = hawb_weight_pair.Key;
                    WeightMap weight_map = hawb_weight_pair.Value;

                    double? weight_sum = 0;
                    double? cost_sum = 0;
                    foreach (var weight_cost_pair in weight_map)
                    {
                        weight_sum += weight_cost_pair.Key;
                        cost_sum += weight_cost_pair.Value;
                    }

                    WeightRangeMap weight_range_map = ref_region_map[region][target_service_level];
                    foreach (KeyValuePair<Weight, double> weight_rate_pair in weight_range_map)
                    {
                        double min_weight = weight_rate_pair.Key.MinWeight;
                        double max_weight = weight_rate_pair.Key.MaxWeight;
                        double rate = weight_rate_pair.Value;

                        if ((weight_sum >= min_weight) && (weight_sum <= max_weight))
                        {
                            double? cost_consolidated = rate * weight_sum;

                            worksheet.Cell(row, col).Value = ship_to_id;
                            ++col;
                            worksheet.Cell(row, col).Value = hawb;
                            ++col;
                            worksheet.Cell(row, col).Value = weight_sum;
                            ++col;
                            worksheet.Cell(row, col).Value = cost_sum;
                            ++col;
                            worksheet.Cell(row, col).Value = cost_consolidated;

                            ++row;
                            col = 1;
                            break;
                        }
                    }
                }
            }

            // Write excel file to disk
            string fullPath = System.IO.Path.Combine(outputFolderPath, outputFile2TextBox.Text);
            workbook.SaveAs($"{fullPath}.xlsx");

            workbook.Dispose();
        }
    }
}