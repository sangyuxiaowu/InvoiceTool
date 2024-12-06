using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Writer;
using ZXing;

namespace 发票助手
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private void Main_Load(object sender, EventArgs e)
        {
            InitializeDataGridView();
        }

        private void InitializeDataGridView()
        {
            // 隐藏存储文件的路径
            dgvPdfFiles.Columns.Add(new DataGridViewColumn()
            {
                HeaderText = "文件路径",
                Name = "FilePath",
                Visible = false,
                CellTemplate = new DataGridViewTextBoxCell()
            });
            dgvPdfFiles.Columns.Add(new DataGridViewColumn()
            {
                HeaderText = "文件名",
                Name = "FileName",
                Width = 200,
                CellTemplate = new DataGridViewTextBoxCell()
            });
            dgvPdfFiles.Columns.Add(new DataGridViewColumn()
            {
                HeaderText = "发票号码",
                Name = "InvoiceNo",
                Width = 150,
                CellTemplate = new DataGridViewTextBoxCell()
            });
            dgvPdfFiles.Columns.Add(new DataGridViewColumn()
            {
                HeaderText = "开票日期",
                Name = "InvoiceDate",
                Width = 80,
                CellTemplate = new DataGridViewTextBoxCell()
            });
            dgvPdfFiles.Columns.Add(new DataGridViewColumn()
            {
                HeaderText = "开票类目",
                Name = "InvoiceType",
                Width = 120,
                CellTemplate = new DataGridViewTextBoxCell()
            });
            // 发票金额
            dgvPdfFiles.Columns.Add("InvoiceAmount", "金额");


            dgvPdfFiles.AllowUserToAddRows = false;
            dgvPdfFiles.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }

        /// <summary>
        /// 正则匹配大写金额
        /// [壹贰叁肆伍陆柒捌玖拾]\s?[零壹贰叁肆伍陆柒捌玖拾佰仟万亿整元圆角分\s]+[整元圆角分]
        /// </summary>
        private static readonly Regex amountRegex = new Regex(@"[壹贰叁肆伍陆柒捌玖拾]\s?[零壹贰叁肆伍陆柒捌玖拾佰仟万亿整元圆角分\s]+[整元圆角分]", RegexOptions.Compiled);

        /// <summary>
        /// 正则匹配年月日
        /// 开票日期[:：]\s*\d{4}年\d{2}月\d{2}日
        /// </summary>
        private static readonly Regex dateRegex = new Regex(@"\d{4}年\d{2}月\d{2}日", RegexOptions.Compiled);

        /// <summary>
        /// 匹配发票号码
        /// </summary>
        private static readonly Regex noRegex = new Regex(@"发票号码[:：]\s*(\d+)", RegexOptions.Compiled);

        /// <summary>
        /// 匹配类目
        /// 匹配到第一个，然后去除两边的*号
        /// </summary>
        private static readonly Regex typeRegex = new Regex(@"\*.*?\*", RegexOptions.Compiled);

        /// <summary>
        /// 匹配金额
        /// </summary>
        private static readonly Regex amountRegex2 = new Regex(@"[¥￥]\s*([0-9]+[.][0-9]{2})", RegexOptions.Compiled );


        /// <summary>
        /// 加载指定目录下的PDF文件，并分析其中的二维码信息
        /// </summary>
        /// <param name="dir"></param>
        /// <param name="isClear"></param>
        private void LoadPdfFiles(string dir,bool isClear = true)
        {
            txtStatus.Text = "正在加载PDF文件...";

            if (isClear)
            {
                dgvPdfFiles.Rows.Clear();
            }
                
            DirectoryInfo dirinfo = new DirectoryInfo(dir);
            FileInfo[] files = dirinfo.GetFiles("*.pdf");

            foreach (FileInfo file in files)
            {
                AnalyzePdfFile(file.FullName, file.Name);
            }

            txtStatus.Text = "PDF文件加载完成";

            // 统计更新
            UpdateStatistics();

        }

        // 处理分析单个PDF文件
        private void AnalyzePdfFile(string fullname,string name)
        {
            using (PdfDocument document = PdfDocument.Open(fullname))
            {
                Page firstPage = document.GetPages().FirstOrDefault();
                if (firstPage != null)
                {
                    // 处理二维码
                    var images = firstPage.GetImages().Where(i => i.HeightInSamples == i.WidthInSamples && i.WidthInSamples > 100 && i.HeightInSamples > 100);
                    var firstImage = images.FirstOrDefault();
                    if (firstImage != null)
                    {
                        var bitmap = ConvertPdfImageToBitmap(firstImage);
                        var qcinfo = bitmap is null ? null : new BarcodeReader().Decode(bitmap);
                        var result = qcinfo is null ? "0,99,?,?,?,?,?" : qcinfo.Text;
                        if (result != null)
                        {
                            string[] values = result.Split(',');
                            if (values.Length > 6)
                            {
                                // 格式化日期统一
                                var invoiceDate = values[5].Contains("-") || values[5] == "?" ? values[5] : values[5].Insert(4, "-").Insert(7, "-");

                                dgvPdfFiles.Rows.Add(fullname, name, values[3], invoiceDate, "", values[4]);
                            }
                            else
                            {
                                dgvPdfFiles.Rows.Add(fullname, name, "?", "?", "?", "?");
                            }

                            // 处理文字内容，解决非31,32数电的发票
                            var text = firstPage.Text;
                            if (values[1] != "31" && values[1] != "32")
                            {
                                var matches = amountRegex2.Matches(text);
                                // 取所有符合条件的金额中最大的
                                decimal amount = 0;
                                foreach (Match match in matches)
                                {
                                    var value = Convert.ToDecimal(match.Groups[1].Value);
                                    if (value > amount)
                                    {
                                        amount = value;
                                    }
                                }
                                dgvPdfFiles.Rows[dgvPdfFiles.Rows.Count - 1].Cells["InvoiceAmount"].Value = amount;
                            }

                            // 处理类目
                            var typeMatch = typeRegex.Match(text);
                            if (typeMatch.Success)
                            {
                                var type = typeMatch.Value.Trim('*');
                                dgvPdfFiles.Rows[dgvPdfFiles.Rows.Count - 1].Cells["InvoiceType"].Value = type;
                            }
                            // 处理发票号码
                            if(dgvPdfFiles.Rows[dgvPdfFiles.Rows.Count - 1].Cells["InvoiceNo"].Value.ToString() == "?")
                            {
                                var noMatch = noRegex.Match(text);
                                if (noMatch.Success)
                                {
                                    var no = noMatch.Groups[1].Value;
                                    dgvPdfFiles.Rows[dgvPdfFiles.Rows.Count - 1].Cells["InvoiceNo"].Value = no;
                                }
                            }
                            // 开票日期
                            if (dgvPdfFiles.Rows[dgvPdfFiles.Rows.Count - 1].Cells["InvoiceDate"].Value.ToString() == "?")
                            {
                                var dateMatch = dateRegex.Match(text);
                                if (dateMatch.Success)
                                {
                                    var date = dateMatch.Value;
                                    // 修改日期格式
                                    date = date.Replace("年", "-").Replace("月", "-").Replace("日", "");
                                    dgvPdfFiles.Rows[dgvPdfFiles.Rows.Count - 1].Cells["InvoiceDate"].Value = date;
                                }
                            }

                        }
                    }

                }
            }
        }


        /// <summary>
        /// 统计更新
        /// </summary>
        private void UpdateStatistics()
        {
            txtN.Text = dgvPdfFiles.Rows.Count.ToString();
            decimal total = 0;
            foreach (DataGridViewRow row in dgvPdfFiles.Rows)
            {
                if (row.Cells["InvoiceAmount"].Value != null)
                {
                    total += Convert.ToDecimal(row.Cells["InvoiceAmount"].Value);
                }
            }
            txtS.Text = total.ToString();

            txtB.Text = UtilTool.ConvertToChinese(total);

            txtStatus.Text = "已更新统计";
        }

        static Bitmap ConvertPdfImageToBitmap(IPdfImage pdfImage)
        {
            pdfImage.TryGetPng(out byte[] pngBytes);
            if (pngBytes == null)
            {
                return null;
            }
            using (MemoryStream ms = new MemoryStream(pngBytes))
            {
                return new Bitmap(ms);
            }
        }

        static string ConvertType(string id)
        {
            switch (id)
            {
                case "01":
                case "08":
                case "31":
                    return "专票";
                case "04":
                case "10":
                case "32":
                    return "普票";
                default:
                    return "未知";
            }
        }

        #region 文件选择

        /// <summary>
        /// 选择文件夹
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSelectFolder_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog()
            {
                Description = "请选择PDF文件所在目录",
                ShowNewFolderButton = false
            })
            {
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    txtDir.Text = fbd.SelectedPath;
                    LoadPdfFiles(fbd.SelectedPath);
                }
            }
        }


        private void txtDir_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files.Length > 0)
            {
                string dir = files[0];
                if (Directory.Exists(dir))
                {
                    txtDir.Text = dir;
                    LoadPdfFiles(dir);
                }
            }

        }

        private void txtDir_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Link;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }

        }

        private void dgvPdfFiles_DragDrop(object sender, DragEventArgs e)
        {
            // 支持文件夹和多个PDF文件
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files.Length > 0)
            {
                foreach (string path in files)
                {
                    if (Directory.Exists(path))
                    {
                        txtDir.Text = path;
                        LoadPdfFiles(path, false);
                    }
                    else
                    {
                        FileInfo file = new FileInfo(path);
                        if (file.Extension == ".pdf")
                        {
                            AnalyzePdfFile(file.FullName, file.Name);
                        }
                    }
                }
                UpdateStatistics();
            }
        }

        private void dgvPdfFiles_DragEnter(object sender, DragEventArgs e)
        {
            
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Link;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        #endregion


        #region 功能处理

        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExport_Click(object sender, EventArgs e)
        {
            // 将列表导出CSV文件

            using (SaveFileDialog sfd = new SaveFileDialog()
            {
                FileName = "发票数据.csv",
                Filter = "CSV文件|*.csv",
                Title = "保存CSV文件"
            })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    string outputFilePath = sfd.FileName;

                    using (StreamWriter sw = new StreamWriter(outputFilePath, false, Encoding.UTF8))
                    {
                        // 带 BOM 的 UTF-8 文件头
                        sw.WriteLine("\uFEFF文件名,发票号码,开票日期,开票类目,金额");

                        foreach (DataGridViewRow row in dgvPdfFiles.Rows)
                        {
                            string fileName = row.Cells["FileName"].Value.ToString();
                            string invoiceNo = row.Cells["InvoiceNo"].Value.ToString();
                            string invoiceDate = row.Cells["InvoiceDate"].Value.ToString();
                            string invoiceType = row.Cells["InvoiceType"].Value.ToString();
                            string invoiceAmount = row.Cells["InvoiceAmount"].Value.ToString();

                            sw.WriteLine($"{fileName},{invoiceNo},{invoiceDate},{invoiceType},{invoiceAmount}");
                        }
                    }

                    // 询问是否打开文件
                    if (MessageBox.Show("CSV文件导出完成，是否打开？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start(outputFilePath);
                    }

                    txtStatus.Text = "CSV文件导出完成";
                }
            }
        }


        /// <summary>
        /// 导出发票类目及金额信息
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExportType_Click(object sender, EventArgs e)
        {
            // 将列表导出CSV文件
            using (SaveFileDialog sfd = new SaveFileDialog()
            {
                FileName = "发票类目金额.csv",
                Filter = "CSV文件|*.csv",
                Title = "保存CSV文件"
            })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    string outputFilePath = sfd.FileName;
                    using (StreamWriter sw = new StreamWriter(outputFilePath, false, Encoding.UTF8))
                    {
                        // 带 BOM 的 UTF-8 文件头
                        sw.WriteLine("\uFEFF开票类目,金额");
                        var query = dgvPdfFiles.Rows.Cast<DataGridViewRow>().GroupBy(r => r.Cells["InvoiceType"].Value.ToString())
                            .Select(g => new
                            {
                                InvoiceType = g.Key,
                                Amount = g.Sum(r => Convert.ToDecimal(r.Cells["InvoiceAmount"].Value))
                            });
                        foreach (var item in query)
                        {
                            sw.WriteLine($"{item.InvoiceType},{item.Amount}");
                        }
                    }

                    // 询问是否打开文件
                    if (MessageBox.Show("CSV文件导出完成，是否打开？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start(outputFilePath);
                    }

                    txtStatus.Text = "CSV文件导出完成";
                }
            }

        }

        /// <summary>
        /// PDF 合并
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnMerge_Click(object sender, EventArgs e)
        {
            // 将列表的PDF文件合并成一个文件

            using (SaveFileDialog sfd = new SaveFileDialog()
            {
                Filter = "PDF文件|*.pdf",
                Title = "保存合并后的PDF文件"
            })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    string outputFilePath = sfd.FileName;
                    string[] pdfFiles = dgvPdfFiles.Rows.Cast<DataGridViewRow>().Select(r => r.Cells["FilePath"].Value.ToString()).ToArray();
                    MergePdfFiles(pdfFiles, outputFilePath);
                    txtStatus.Text = "PDF文件合并完成";
                }
            }
        }

        // 合并PDF文件
        private void MergePdfFiles(string[] pdfFiles, string outputFilePath)
        {
            PdfDocumentBuilder builder = new PdfDocumentBuilder();

            foreach (string pdfFile in pdfFiles)
            {
                using (PdfDocument inputDocument = PdfDocument.Open(pdfFile))
                {
                    for (var i = 0; i < inputDocument.NumberOfPages; i++)
                    {
                        builder.AddPage(inputDocument, i + 1);
                    }
                }
            }

            //保存PDF文件
            var documentBytes = builder.Build();
            File.WriteAllBytes(outputFilePath, documentBytes);
        }

        /// <summary>
        /// 重置
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnReset_Click(object sender, EventArgs e)
        {
            dgvPdfFiles.Rows.Clear();
            txtDir.Text = "";
            txtN.Text = "0";
            txtS.Text = "0";
            txtB.Text = "";
        }

        /// <summary>
        /// 打印
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPrint_Click(object sender, EventArgs e)
        {
            // 在临时目录下合并PDF文件，并打印
            string tempDir = Path.Combine(Path.GetTempPath(), "发票助手");
            if (!Directory.Exists(tempDir))
            {
                Directory.CreateDirectory(tempDir);
            }

            string tempPdfFile = Path.Combine(tempDir, "temp.pdf");
            string[] pdfFiles = dgvPdfFiles.Rows.Cast<DataGridViewRow>().Select(r => r.Cells["FilePath"].Value.ToString()).ToArray();
            MergePdfFiles(pdfFiles, tempPdfFile);

            // 打印
            PrintPdfFile(tempPdfFile);
        }

        /// <summary>
        /// 打印指定文件
        /// </summary>
        /// <param name="tempPdfFile"></param>
        private async void PrintPdfFile(string tempPdfFile)
        {
            System.Diagnostics.Process.Start("explorer", tempPdfFile);
            await Task.Delay(1000);
            // 发送 Ctrl + P
            SendKeys.SendWait("^(p)");
        }



        #endregion

        
    }
}