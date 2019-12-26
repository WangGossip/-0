using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using NPOI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using MSWord = Microsoft.Office.Interop.Word;
using NPOI.OpenXmlFormats.Wordprocessing;
using System.Reflection;
using WpsApiEx;
//using Aspose.Words;

namespace WindowsForms_Word_To_Excel
{
    public partial class Form1 : Form
    {
        private string input_filename, outputfile_foilder, tmp_info_file, retire_time_file;
        bool use_wps = false;
        int retire_year1, retire_year2, retire_month;
        object doc_standard, tmp_info_name;
        private List<string> FirstClassName = new List<string>();
        private List<string> output_filenames = new List<string>();
        private List<string> class_names = new List<string>();
        private List<string> retire_month_list, retire_year_list, retire_list1, retire_list2, submit_year_list, submit_list1, submit_list2;
        private string[] info_names_new = { "到龄月份", "姓名", "一级单位" };
        private int[] info_column_num;
        private string[] info_names = { "姓名", "一级单位", "到达退休年龄时间", "领退休金时间", "递交时间", "申报备案时间", "最后署名时间" };
        private string[] bookmark_names = { "Name", "FirstClass", "RetireTime1", "RetireTime2", "SubmitTime1", "SubmitTime2", "DeclareTime1", "DeclareTime2", "LastSignTime" };
        string[] info_tmp = new string[7];
        string[] info_use = new string[9];
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            input_filename = "";
            //filename_end = "退休告知";
            outputfile_foilder = System.AppDomain.CurrentDomain.BaseDirectory;
            label7.Text = outputfile_foilder;
            tmp_info_file = "tmp.docx";
            tmp_info_name = tmp_info_file;
            doc_standard = "";
            info_column_num = new int[info_names_new.Length];
            retire_year_list = new List<string>();
            retire_list1 = new List<string>();
            retire_list2 = new List<string>();
            submit_year_list = new List<string>();
            submit_list1 = new List<string>();
            submit_list2 = new List<string>();
            retire_month_list = new List<string>();
            init_num();
        }

        private void 导入文件ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            /*选择并打开文件，读取数据，关闭文件*/
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Title = "请选择一个符合格式的excel表格";
            dialog.Filter = "Excel文件(*.xlsx)|*.xlsx";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                input_filename = dialog.FileName;
            }
            label1.Text = input_filename;
        }

        private void 选择保存路径ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
            dialog.Description = "请选择目标文件夹";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (string.IsNullOrEmpty(dialog.SelectedPath))
                {
                    MessageBox.Show(this, "文件夹路径不能为空", "提示");
                    return;
                }
            }
            outputfile_foilder = dialog.SelectedPath;
            label7.Text = outputfile_foilder;
        }

        private void 选择文件模板ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            /*选择并打开文件，读取数据，关闭文件*/
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Title = "请选择一个符合格式的word文件";
            dialog.Filter = "Excel文件(*.docx)|*.docx";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                doc_standard = dialog.FileName;
            }
            label9.Text = doc_standard.ToString();
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            //当用户按下的键盘的键不在数字位的话，就禁止输入
            if (e.KeyChar != 8 && !Char.IsDigit(e.KeyChar))//如果不是输入数字就不让输入
            {
                e.Handled = true;
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            int iMax = 9999;//首先设置上限值
            if (textBox3.Text != null && textBox3.Text != "")//判断TextBox的内容不为空，如果不判断会导致后面的非数字对比异常
            {
                if (int.Parse(textBox1.Text) > iMax)//num就是传进来的值,如果大于上限（输入的值），那就强制为上限-1，或者就是上限值？
                {
                    textBox3.Text = (iMax - 1).ToString();
                }
            }
        }

        private void label18_Click(object sender, EventArgs e)
        {

        }

        private void 使用wps格式ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (use_wps)
            {
                use_wps = false;
                MessageBox.Show("当前不使用wps格式！");
                使用wps格式ToolStripMenuItem.Text = "使用wps格式";
            }
            else
            {
                use_wps = true;
                MessageBox.Show("当前使用wps格式！");
                使用wps格式ToolStripMenuItem.Text = "不使用wps格式";
            }
        }

        private void 使用帮助ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (textBox17.Text == "")
            {
                textBox17.Text = "1、首先选择输入文件（表格）\r\n2、之后选择退休告知时间表格\r\n3、之后选择保存" +
                    "路径并输入文件名（输入名称即可，不需要添加后缀）\r\n4、输入相关的年、月、日信息（仅能输入数字" +
                    "且有输入限制，生成的文件中这些信息相同）\r\n5、点击“开始”按钮，当提示“处理完成”时可以退出程序\n";
            }
            else
                textBox17.Text = "";
            return;
        }

        /// <summary>
        /// 从模板文件（一个word）
        /// </summary>
        /// <param name="sorceDocPath">源文件路径</param>
        /// <returns></returns>
        protected MSWord.Document copyWordDoc(object sorceDocPath)
        {
            MSWord.Application wordApp;//Word应用程序变量 
            MSWord.Document newWordDoc;//Word文档变量
            object readOnly = false;
            object isVisible = false;
            //初始化
            //由于使用的是COM库，因此有许多变量需要用Missing.Value代替
            wordApp = new MSWord.Application();
            Object Nothing = System.Reflection.Missing.Value;
            newWordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            MSWord.Document openWord;
            openWord = wordApp.Documents.Open(ref sorceDocPath, ref Nothing, ref readOnly, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref isVisible, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            openWord.Select();
            openWord.Sections[1].Range.Copy();
            object start = 0;
            MSWord.Range newRang = newWordDoc.Range(ref start, ref start);
            newWordDoc.Sections[1].Range.PasteAndFormat(MSWord.WdRecoveryType.wdPasteDefault);
            openWord.Close(ref Nothing, ref Nothing, ref Nothing);
            /*释放内存十分重要*/
            object IsSave = true;
            object missing = System.Reflection.Missing.Value;
            wordApp.Quit(ref IsSave, ref missing, ref missing);//退出程序，相当于关闭word   
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp); //释放内存
            wordApp = null;//内存释放完成后，切记要将wordApp 置为空，很重要！
            return newWordDoc;
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            //当用户按下的键盘的键不在数字位的话，就禁止输入
            if (e.KeyChar != 8 && !Char.IsDigit(e.KeyChar))//如果不是输入数字就不让输入
            {
                e.Handled = true;
            }
        }

        /*处理从excel中获取的字符串*/
        protected void use_str_info(string month)
        {
            int retire_month2, month_index;
            month_index = retire_month_list.IndexOf(month);//获取索引
            if (retire_month == 12)
            {
                retire_year2 = retire_year1 + 1;
                retire_month2 = 1;
            }
            else
            {
                retire_year2 = retire_year1;
                retire_month2 = retire_month + 1;
            }
            //到龄时间1
            info_use[2] = retire_year1 + "年" + retire_month + "月";
            //到龄时间2
            info_use[3] = retire_year2 + "年" + retire_month2 + "月";
            //递交时间1
            info_use[4] = retire_year_list[month_index] + retire_list1[month_index];
            //递交时间2
            info_use[5] = retire_list2[month_index];
            //备案时间1
            info_use[6] = submit_year_list[month_index] + submit_list1[month_index];
            //备案时间2
            info_use[7] = submit_list2[month_index];
            return;
            //for (int i = 0; i < 4; i++)
            //    info_use[i] = info_tmp[i];
            //info_use[4] = info_tmp[4].Substring(0, 10);//SubmitTime1
            //info_use[5] = info_tmp[4].Substring(12);//SubmitTime2
            //info_use[6] = info_tmp[5].Substring(0, 10);//DeclareTime1
            //info_use[7] = info_tmp[5].Substring(12);//DeclareTime2
            //info_use[8] = info_tmp[6];//LastSignTime
        }

        /// <summary>
        /// 读取相关信息存入list中
        /// </summary>
        protected void Read_retire_time()
        {
            int line_num, i;
            IWorkbook wk;
            FileStream fs;
            IRow row;
            ISheet sheet;
            /*打开文件，读取表格*/
            fs = File.OpenRead(retire_time_file);
            if (fs == null)
            {
                MessageBox.Show("文件打开失败，请重新选择文件");
                return;
            }
            wk = new XSSFWorkbook(fs);
            fs.Close();
            sheet = wk.GetSheetAt(0);
            line_num = sheet.LastRowNum + 1;//总行数
            for (i = 1; i < line_num; i++)
            {
                row = sheet.GetRow(i);
                if (row.GetCell(0).ToString() == "")
                    break;
                retire_month_list.Add(row.GetCell(0).ToString());//月份
                retire_year_list.Add(row.GetCell(1).ToString());//退休年份
                retire_list1.Add(row.GetCell(2).ToString());//退休时间1
                retire_list2.Add(row.GetCell(3).ToString());//退休时间2
                submit_year_list.Add(row.GetCell(4).ToString());//提交年份
                submit_list1.Add(row.GetCell(5).ToString());//提交时间1
                submit_list2.Add(row.GetCell(6).ToString());//提交时间2
            }
            return;
        }

        private void 选择退休告知时间ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            /*选择并打开文件，读取数据，关闭文件*/
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Title = "请选择一个符合格式的excel表格";
            dialog.Filter = "Excel文件(*.xlsx)|*.xlsx";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                retire_time_file = dialog.FileName;
            }
            label31.Text = retire_time_file;
        }

        protected void Delete_tmp()
        {
            if (File.Exists(tmp_info_file))
                File.Delete(tmp_info_file);
            return;
        }

        /*判断是否所有信息都已经输入了*/
        protected bool Judge_Info_Full()
        {
            bool info_full = false;
            /*判断模板文件是否选择*/
            if (doc_standard.ToString() == "")
            {
                MessageBox.Show("请先选择模板文件!");
            }
            /*判断是否已经打开文件*/
            else if (input_filename == "")
            {
                MessageBox.Show("请先选择文件!");
            }
            else if (retire_time_file == "")
            {
                MessageBox.Show("请选择退休告知表格");
            }
            else if (textBox3.Text == "")
            {
                MessageBox.Show("请先输入到龄年份");
            }
            else if (textBox14.Text == "" || textBox15.Text == "" || textBox16.Text == "")
            {
                MessageBox.Show("请先输入完整的最后署名时间");
            }
            else if (textBox2.Text == "")
            {
                MessageBox.Show("请先输入文件名");
            }
            else
            {
                info_full = true;
            }
            return info_full;
        }

        /// <summary>
        /// 初始化标记数组值为-1
        /// </summary>
        protected void init_num()
        {
            for (int i = 0; i < info_column_num.Length; i++)
                info_column_num[i] = -1;
            return;
        }
        /// <summary>
        /// 判断excel文件是否符合格式（包含对应字符串）,并给索引数组对应赋值
        /// </summary>
        /// <param name="sheet">表格参数</param>
        /// <returns></returns>
        protected bool Jduge_Excel_Column(ISheet sheet)
        {
            /*变量区*/
            bool Excel_OK = true;
            int i, j, column_num, info_excel_len;
            string value_tmp;
            IRow row = sheet.GetRow(0);
            column_num = row.LastCellNum;//总列数，实际为固定值
            info_excel_len = info_names_new.Length;
            for (i = 0; i < info_excel_len; i++)
            {
                for (j = 0; j < column_num; j++)
                {
                    value_tmp = row.GetCell(j).ToString();
                    if (string.Compare(value_tmp, info_names_new[i]) == 0)
                    {
                        info_column_num[i] = j;//数组元素代表了实际第几列
                        break;
                    }
                }
                if (j == column_num && info_column_num[i] == -1)//说明所有列中没有需要的内容
                {
                    Excel_OK = false;
                    break;
                }
            }
            return Excel_OK;
        }

        protected void Sta_Class_Init(ISheet sheet)
        {
            int i, column_num, class_num, line_num;
            IRow row = sheet.GetRow(0);
            line_num = sheet.LastRowNum + 1;//总行数
            column_num = row.LastCellNum;//总列数，实际为固定值
            class_num = 8;
            for (i = 0; i < column_num; i++)
            {
                if (row.GetCell(i).ToString() == info_names_new[2])
                {
                    class_num = i;
                    break;
                }
            }
            for (i = 1; i < line_num; i++)
            {
                row = sheet.GetRow(i);
                if (row.GetCell(class_num).ToString() == "")
                    break;
                class_names.Add(row.GetCell(class_num).ToString());
            }
            class_names = class_names.Distinct().ToList();//去除重复元素
            return;
        }

        private void 开始ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!use_wps)
            {
                /*判断各类信息输入情况*/
                if (!Judge_Info_Full())
                    return;
                /*变量区域*/
                #region 变量
                FileStream fs;
                IWorkbook wk;
                int[] class_nums;
                int i, j, line_num, column_num, use_info_num;//计数
                object out_file_path, bkObj;
                object readOnly = false;
                object isVisible = false;
                object start = 0;
                Object Nothing = Missing.Value;
                object unite = MSWord.WdUnits.wdStory;
                MSWord.Application wordApp;                   //Word应用程序变量 
                MSWord.Document wordDoc_tmp, openWord, newWordDoc;  //Word文档变量
                                                                    //XWPFDocument doc = new XWPFDocument();
                #endregion

                wordApp = new MSWord.Application(); //初始化
                                                    //由于使用的是COM库，因此有许多变量需要用Missing.Value代替
                                                    /*判断是否符合格式，仅检查第一行*/
                fs = File.OpenRead(input_filename);
                if (fs == null)
                {
                    MessageBox.Show("文件打开失败，请重新选择文件");
                    return;
                }
                wk = new XSSFWorkbook(fs);
                fs.Close();
                Read_retire_time();
                ISheet sheet = wk.GetSheetAt(0);
                line_num = sheet.LastRowNum + 1;//总行数
                IRow row = sheet.GetRow(0);  //读取当前行数据
                use_info_num = info_names_new.Length;//用到的列数
                column_num = row.LastCellNum;//总列数，实际为固定值
                if (!Jduge_Excel_Column(sheet))
                {
                    MessageBox.Show("文件格式不符合!请修改后重试");
                    return;
                }
                /*判断文件名是否有问题*/
                out_file_path = outputfile_foilder + "\\" + textBox2.Text.ToString() + ".pdf";//注意组合后的路径名称
                if (File.Exists(out_file_path.ToString()))
                {
                    MessageBox.Show("输入的文件已经存在，请重新输入！");
                    return;
                }
                /*获取所有的学院名称*/
                Sta_Class_Init(sheet);
                //初始化计数的数组
                class_nums = new int[class_names.Count];
                for (i = 0; i < class_names.Count; i++)
                {
                    class_nums[i] = 0;
                }
                /*获取字符串中信息*/
                //到龄年份
                retire_year1 = Convert.ToInt32(textBox3.Text.ToString());//到龄年份
                ////递交时间1
                //info_use[4] = textBox4.Text.ToString() + "年" + textBox5.Text.ToString() + "月" + textBox6.Text.ToString() + "日";
                ////递交时间2
                //info_use[5] = textBox7.Text.ToString() + "月" + textBox8.Text.ToString() + "日";
                ////申报备案时间1
                //info_use[6] = textBox13.Text.ToString() + "年" + textBox12.Text.ToString() + "月" + textBox11.Text.ToString() + "日";
                ////申报备案时间2
                //info_use[7] = textBox10.Text.ToString() + "月" + textBox9.Text.ToString() + "日";
                //最后署名时间
                info_use[8] = textBox16.Text.ToString() + "年" + textBox15.Text.ToString() + "月" + textBox14.Text.ToString() + "日";

                //先定义一个外面的文件
                /*循环遍历每一行，写入一个文件即可*/
                newWordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);//初始化输出文件
                label2.Text = line_num.ToString();
                /*循环体*/
                for (i = 1; i < line_num; i++)
                {
                    /*变量区*/
                    //string name_now, workplace_now, retire_year, retire_money_time, submit_time, declare_time, last_sign_time;
                    #region 处理读取的字符串
                    row = sheet.GetRow(i);
                    //姓名
                    info_use[0] = row.GetCell(info_column_num[1]).ToString();
                    if (info_use[0] == "")//判断是否存在意外空输入情况
                        break;
                    label2.Text = "正在处理第" + i + "条信息……";
                    //一级单位
                    info_use[1] = row.GetCell(info_column_num[2]).ToString();
                    class_nums[class_names.IndexOf(info_use[1])]++;
                    retire_month = Convert.ToInt32(row.GetCell(info_column_num[0]).ToString());//到龄月份
                                                                                               /*日志信息*/
                    textBox1.Text += "正在处理第" + i + "条信息" + "\r\n" + info_use[0] + " " + info_use[1] + "\r\n" + "……" + "\r\n";
                    use_str_info(row.GetCell(info_column_num[0]).ToString());//处理字符串到需要的格式
                    #endregion
                    #region 使用aspore.word 赋值模板内容
                    //Document std_doc = new Document(doc_standard.ToString()); //模板文件
                    //Document tmp_doc = std_doc;//临时文件
                    //DocumentBuilder builder = new DocumentBuilder(tmp_doc);   //操作word
                    #endregion
                    /*先复制到tmp文件中，再修改tmp中内容并复制，粘贴到新的文件中*/
                    #region 复制模板中内容到临时文件
                    wordDoc_tmp = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);//初始化中间文件
                    openWord = wordApp.Documents.Open(ref doc_standard, ref Nothing, ref readOnly, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref isVisible, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
                    openWord.Select();//选中openword
                    openWord.Sections[1].Range.Copy();
                    MSWord.Range newRang = wordDoc_tmp.Range(ref start, ref start);//改为当前操作的文档
                    wordDoc_tmp.Sections[1].Range.PasteAndFormat(MSWord.WdRecoveryType.wdPasteDefault);
                    //导出到tmp文件中
                    wordDoc_tmp.SaveAs(ref tmp_info_name, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
                    openWord.Close();
                    #endregion
                    #region 修改文件内容，并复制到输出文件
                    //打开临时文件，根据内容进行修改
                    wordDoc_tmp.Select();
                    for (j = 0; j < 9; j++)//遍历书签进行插入
                    {
                        bkObj = bookmark_names[j];//书签的字符串转成对应的object型
                        wordDoc_tmp.Bookmarks.get_Item(ref bkObj).Select();//选中书签对应的位置
                        wordApp.Selection.TypeText(info_use[j]);
                    }
                    wordDoc_tmp.Sections[1].Range.Copy();//复制临时文件内容
                    newWordDoc.Select();//选中文件
                    wordApp.Selection.EndKey(ref unite, ref Nothing);//将光标移到文本末尾
                    newWordDoc.Paragraphs.Last.Range.PasteAndFormat(MSWord.WdRecoveryType.wdPasteDefault);//在文件末尾插入复制后的文档
                    wordDoc_tmp.Close();//关闭文件
                    #endregion
                    Delete_tmp();//删除临时文件
                    textBox1.Text += "第" + i + "条信息处理完成" + "\r\n";
                }
                //循环结束后保存以及关闭文件
                //WdSaveFormat为Word 默认文档的保存格式
                object format = MSWord.WdSaveFormat.wdFormatPDF;
                //object format = MSWord.WdSaveFormat.wdFormatDocumentDefault;// office 2007就是wdFormatDocumentDefault
                newWordDoc.SaveAs(ref out_file_path, ref format, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
                newWordDoc.Close();
                /*释放内存十分重要*/
                object IsSave = true;
                object missing = System.Reflection.Missing.Value;
                wordApp.Quit(ref IsSave, ref missing, ref missing);//退出程序，相当于关闭word   
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp); //释放内存
                wordApp = null;//内存释放完成后，切记要将wordApp 置为空，很重要！
                MessageBox.Show("处理完成！");
                textBox1.Text = "";
                for (i = 0; i < class_names.Count; i++)
                {
                    textBox1.Text += class_names[i] + "\r\n" + class_nums[i] + "\r\n";
                }
                label2.Text = "处理完成";

            }
            else
            {
                /*判断各类信息输入情况*/
                if (!Judge_Info_Full())
                    return;
                /*变量区域*/
                #region 变量
                FileStream fs;
                IWorkbook wk;
                int[] class_nums;
                int i, j, line_num, column_num, use_info_num;//计数
                object out_file_path, bkObj;
                object readOnly = false;
                object isVisible = false;
                object start = 0;
                Object Nothing = Missing.Value;
                object unite = Word.WdUnits.wdStory;
                Word.Application wordApp;                   //Word应用程序变量 
                Word.Document wordDoc_tmp, openWord, newWordDoc;  //Word文档变量
                                                                    //XWPFDocument doc = new XWPFDocument();
                #endregion

                wordApp = new Word.Application(); //初始化
                                                    //由于使用的是COM库，因此有许多变量需要用Missing.Value代替
                                                    /*判断是否符合格式，仅检查第一行*/
                fs = File.OpenRead(input_filename);
                if (fs == null)
                {
                    MessageBox.Show("文件打开失败，请重新选择文件");
                    return;
                }
                wk = new XSSFWorkbook(fs);
                fs.Close();
                Read_retire_time();
                ISheet sheet = wk.GetSheetAt(0);
                line_num = sheet.LastRowNum + 1;//总行数
                IRow row = sheet.GetRow(0);  //读取当前行数据
                use_info_num = info_names_new.Length;//用到的列数
                column_num = row.LastCellNum;//总列数，实际为固定值
                if (!Jduge_Excel_Column(sheet))
                {
                    MessageBox.Show("文件格式不符合!请修改后重试");
                    return;
                }
                /*判断文件名是否有问题*/
                out_file_path = outputfile_foilder + "\\" + textBox2.Text.ToString() + ".pdf";//注意组合后的路径名称
                if (File.Exists(out_file_path.ToString()))
                {
                    MessageBox.Show("输入的文件已经存在，请重新输入！");
                    return;
                }
                /*获取所有的学院名称*/
                Sta_Class_Init(sheet);
                //初始化计数的数组
                class_nums = new int[class_names.Count];
                for (i = 0; i < class_names.Count; i++)
                {
                    class_nums[i] = 0;
                }
                /*获取字符串中信息*/
                //到龄年份
                retire_year1 = Convert.ToInt32(textBox3.Text.ToString());//到龄年份
                ////递交时间1
                //info_use[4] = textBox4.Text.ToString() + "年" + textBox5.Text.ToString() + "月" + textBox6.Text.ToString() + "日";
                ////递交时间2
                //info_use[5] = textBox7.Text.ToString() + "月" + textBox8.Text.ToString() + "日";
                ////申报备案时间1
                //info_use[6] = textBox13.Text.ToString() + "年" + textBox12.Text.ToString() + "月" + textBox11.Text.ToString() + "日";
                ////申报备案时间2
                //info_use[7] = textBox10.Text.ToString() + "月" + textBox9.Text.ToString() + "日";
                //最后署名时间
                info_use[8] = textBox16.Text.ToString() + "年" + textBox15.Text.ToString() + "月" + textBox14.Text.ToString() + "日";

                //先定义一个外面的文件
                /*循环遍历每一行，写入一个文件即可*/
                newWordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);//初始化输出文件
                label2.Text = line_num.ToString();
                /*循环体*/
                for (i = 1; i < line_num; i++)
                {
                    /*变量区*/
                    //string name_now, workplace_now, retire_year, retire_money_time, submit_time, declare_time, last_sign_time;
                    #region 处理读取的字符串
                    row = sheet.GetRow(i);
                    //姓名
                    info_use[0] = row.GetCell(info_column_num[1]).ToString();
                    if (info_use[0] == "")//判断是否存在意外空输入情况
                        break;
                    label2.Text = "正在处理第" + i + "条信息……";
                    //一级单位
                    info_use[1] = row.GetCell(info_column_num[2]).ToString();
                    class_nums[class_names.IndexOf(info_use[1])]++;
                    retire_month = Convert.ToInt32(row.GetCell(info_column_num[0]).ToString());//到龄月份
                                                                                               /*日志信息*/
                    textBox1.Text += "正在处理第" + i + "条信息" + "\r\n" + info_use[0] + " " + info_use[1] + "\r\n" + "……" + "\r\n";
                    use_str_info(row.GetCell(info_column_num[0]).ToString());//处理字符串到需要的格式
                    #endregion
                    #region 使用aspore.word 赋值模板内容
                    //Document std_doc = new Document(doc_standard.ToString()); //模板文件
                    //Document tmp_doc = std_doc;//临时文件
                    //DocumentBuilder builder = new DocumentBuilder(tmp_doc);   //操作word
                    #endregion
                    /*先复制到tmp文件中，再修改tmp中内容并复制，粘贴到新的文件中*/
                    #region 复制模板中内容到临时文件
                    wordDoc_tmp = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);//初始化中间文件
                    openWord = wordApp.Documents.Open(ref doc_standard, ref Nothing, ref readOnly, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref isVisible, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
                    openWord.Select();//选中openword
                    openWord.Sections[1].Range.Copy();
                    Word.Range newRang = wordDoc_tmp.Range(ref start, ref start);//改为当前操作的文档
                    wordDoc_tmp.Sections[1].Range.PasteAndFormat(Word.WdRecoveryType.wdPasteDefault);
                    //导出到tmp文件中
                    wordDoc_tmp.SaveAs(ref tmp_info_name, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
                    openWord.Close();
                    #endregion
                    #region 修改文件内容，并复制到输出文件
                    //打开临时文件，根据内容进行修改
                    wordDoc_tmp.Select();
                    for (j = 0; j < 9; j++)//遍历书签进行插入
                    {
                        bkObj = bookmark_names[j];//书签的字符串转成对应的object型
                        wordDoc_tmp.Bookmarks.get_Item(ref bkObj).Select();//选中书签对应的位置
                        wordApp.Selection.TypeText(info_use[j]);
                    }
                    wordDoc_tmp.Sections[1].Range.Copy();//复制临时文件内容
                    newWordDoc.Select();//选中文件
                    wordApp.Selection.EndKey(ref unite, ref Nothing);//将光标移到文本末尾
                    newWordDoc.Paragraphs.Last.Range.PasteAndFormat(Word.WdRecoveryType.wdPasteDefault);//在文件末尾插入复制后的文档
                    wordDoc_tmp.Close();//关闭文件
                    #endregion
                    Delete_tmp();//删除临时文件
                    textBox1.Text += "第" + i + "条信息处理完成" + "\r\n";
                }
                //循环结束后保存以及关闭文件
                //WdSaveFormat为Word 默认文档的保存格式
                object format = Word.WdSaveFormat.wdFormatPDF;
                //object format = Word.WdSaveFormat.wdFormatDocumentDefault;// office 2007就是wdFormatDocumentDefault
                newWordDoc.SaveAs(ref out_file_path, ref format, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
                newWordDoc.Close();
                /*释放内存十分重要*/
                object IsSave = true;
                object missing = System.Reflection.Missing.Value;
                wordApp.Quit(ref IsSave, ref missing, ref missing);//退出程序，相当于关闭word   
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp); //释放内存
                wordApp = null;//内存释放完成后，切记要将wordApp 置为空，很重要！
                MessageBox.Show("处理完成！");
                textBox1.Text = "";
                for (i = 0; i < class_names.Count; i++)
                {
                    textBox1.Text += class_names[i] + "\r\n" + class_nums[i] + "\r\n";
                }
                label2.Text = "处理完成";

            }

        }
    }
}
