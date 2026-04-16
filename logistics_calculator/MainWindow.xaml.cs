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

namespace logistics_calculator
{
    using Dict4 = System.Collections.Generic.Dictionary<Weight, double>;
    using Dict3 = System.Collections.Generic.Dictionary<string, System.Collections.Generic.Dictionary<Weight, double>>;
    using Dict2 = System.Collections.Generic.Dictionary<string, System.Collections.Generic.Dictionary<string, System.Collections.Generic.Dictionary<Weight, double>>>;
    using Dict1 = System.Collections.Generic.Dictionary<string, System.Collections.Generic.Dictionary<string, System.Collections.Generic.Dictionary<string, System.Collections.Generic.Dictionary<Weight, double>>>>;
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ProcessRefFile(object sender, RoutedEventArgs e)
        {
            string file_path = null;
            OpenFileDialog dialog = new OpenFileDialog();

            if (dialog.ShowDialog() == false)
            {
                return;
            }

            file_path = dialog.FileName;

            FileStream stream = File.Open(file_path, FileMode.Open, FileAccess.Read);
            var reader = ExcelReaderFactory.CreateReader(stream);

            #region get header
            // Get header first            
            Dictionary<string, int> header_index_map = new Dictionary<string, int>();
            header_index_map["DestCountry/RegionCode"] = 0;
            header_index_map["ServiceLevelShortDescription"] = 0;
            header_index_map["PackageTypeCode"] = 0;
            header_index_map["MinWeight"] = 0;
            header_index_map["MaxWeight"] = 0;
            header_index_map["Rate"] = 0;

            int found = 0;
            bool all_correct = false;
            while (reader.Read())
            {
                int column = 0;
                while (true)
                {
                    var value = reader.GetValue(column);
                    if (value == null)
                    {
                        break;
                    }
                    string value_string = value.ToString();
                    if (string.IsNullOrWhiteSpace(value_string))
                    {
                        break;
                    }

                    // Use dictionary key to update value
                    if (header_index_map.ContainsKey(value_string))
                    {
                        header_index_map[value_string] = column;
                        ++found;
                    }

                    if (found == header_index_map.Count)
                    {
                        all_correct = true;
                        break;
                    }

                    ++column;
                }
                if (found > 0)
                {
                    break;
                }
            }
            if (!all_correct)
            {
                return;
            }
            #endregion

            Dict1 dict1 = new Dict1();

            bool abort = false;
            while (reader.Read())
            {
                string region = reader.GetValue(header_index_map["DestCountry/RegionCode"])?.ToString() ?? "";
                string service_level = reader.GetValue(header_index_map["ServiceLevelShortDescription"])?.ToString() ?? "";
                string package_type = reader.GetValue(header_index_map["PackageTypeCode"])?.ToString() ?? "";
                string min_weight_str = reader.GetValue(header_index_map["MinWeight"])?.ToString() ?? "";
                string max_weight_str = reader.GetValue(header_index_map["MaxWeight"])?.ToString() ?? "";
                string rate_str = reader.GetValue(header_index_map["Rate"])?.ToString() ?? "";

                region = region.Replace(" ", "");
                service_level = service_level.Replace(" ", "");
                package_type = package_type.Replace(" ", "");
                min_weight_str = min_weight_str.Replace(" ", "");
                max_weight_str = max_weight_str.Replace(" ", "");
                rate_str = rate_str.Replace(" ", "");

                Weight weight = new Weight();
                if (double.TryParse(min_weight_str, out double min_weight))
                {
                    if (double.TryParse(max_weight_str, out double max_weight))
                    {
                        weight.MinWeight = min_weight;
                        weight.MaxWeight = max_weight;
                    }
                    else
                    {
                        abort = true;
                        break;
                    }
                }
                else
                {
                    abort = true;
                    break;
                }

                if (!double.TryParse(rate_str, out double rate))
                {
                    abort = true;
                    break;
                }

                if (!dict1.ContainsKey(region))
                {
                    dict1[region] = new Dict2();
                }

                if (!dict1[region].ContainsKey(service_level))
                {
                    dict1[region][service_level] = new Dict3();
                }

                if (!dict1[region][service_level].ContainsKey(package_type))
                {
                    dict1[region][service_level][package_type] = new Dict4();
                }

                if (!dict1[region][service_level][package_type].ContainsKey(weight))
                {
                    dict1[region][service_level][package_type] = new Dict4();
                }

                if (dict1[region][service_level][package_type][weight] != rate)
                {
                    dict1[region][service_level][package_type][weight] = rate;
                }
            }

            if (!abort)
            {
                return;
            }

            reader.Close();
            stream.Close();
        }

        private void ProcessQuesFile(object sender, RoutedEventArgs e)
        {

        }

        private void Start(object sender, RoutedEventArgs e)
        {

        }
    }
}