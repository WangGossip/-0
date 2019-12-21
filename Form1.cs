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

namespace WindowsForms_Word_To_Excel
{
    public partial class Form1 : Form
    {
        private string input_filename, filename_end, outputfile_foilder, tmp_info_file;
        object doc_standard, tmp_info_name;
        private List<string> FirstClassName = new List<string>();
        private List<string> output_filenames = new List<string>();
        private string[] info_names = { "姓名", "一级单位", "到达退休年龄时间", "领退休金时间", "递交时间", "申报备案时间", "最后署名时间" };
        private string[] bookmark_names = { "Name","FirstClass", "RetireTime1", "RetireTime2", "SubmitTime1", "SubmitTime2", "DeclareTime1", "DeclareTime2","LastSignTime" };
        string[] info_tmp = new string[7];
        string[] info_use = new string[9];
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            input_filename = "";
            filename_end = "退休告知";
            outputfile_foilder=System.AppDomain.CurrentDomain.BaseDirectory;
            label7.Text = outputfile_foilder;
            tmp_info_file = "tmp.docx";
            tmp_info_name = tmp_info_file;
            doc_standard = "";
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
            return newWordDoc;
        }
        
        /*处理从excel中获取的字符串*/
        protected void use_str_info()
        {
            for (int i = 0; i < 4; i++)
                info_use[i] = info_tmp[i];
            info_use[4] = info_tmp[4].Substring(0, 10);//SubmitTime1
            info_use[5] = info_tmp[4].Substring(12);//SubmitTime2
            info_use[6] = info_tmp[5].Substring(0, 10);//DeclareTime1
            info_use[7] = info_tmp[5].Substring(12);//DeclareTime2
            info_use[8] = info_tmp[6];//LastSignTime
        }
        
        protected void Delete_tmp()
        {
            if (File.Exists(tmp_info_file))
                File.Delete(tmp_info_file);
            return;
        }

        private void 开始ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            /*变量区域*/
            #region 变量
            FileStream fs;
            IWorkbook wk;
            int i, j, line_num, column_num, table_column, table_row;//计数
            bool file_ok = true;
            string value_tmp, word_text;
            object out_file_path, bkObj;
            object readOnly = false;
            object isVisible = false;
            object start = 0;
            Object Nothing = Missing.Value;
            object unite = MSWord.WdUnits.wdStory;
            MSWord.Application wordApp;                   //Word应用程序变量 
            MSWord.Document wordDoc_tmp,openWord,newWordDoc;  //Word文档变量
            XWPFDocument doc = new XWPFDocument();
            #endregion

            wordApp = new MSWord.Application(); //初始化
            //由于使用的是COM库，因此有许多变量需要用Missing.Value代替
            /*判断模板文件是否选择*/
            if (doc_standard.ToString()=="")
            {
                MessageBox.Show("请先选择模板文件!");
                return;
            }
            /*先判断是否已经打开文件*/
            if (input_filename == "")
            {
                MessageBox.Show("请先选择文件!");
                return;
            }
            /*判断是否符合格式，仅检查第一行*/
            fs = File.OpenRead(input_filename);
            if (fs == null)
            {
                MessageBox.Show("文件打开失败，请重新选择文件");
                return;
            }
            wk = new XSSFWorkbook(fs);
            fs.Close();
            ISheet sheet = wk.GetSheetAt(0);
            line_num = sheet.LastRowNum + 1;//总行数
            IRow row = sheet.GetRow(0);  //读取当前行数据
            column_num = row.LastCellNum;//总列数，实际为固定值
            if (column_num != info_names.Length)
            {
                MessageBox.Show("文件格式不符合!请修改后重试");
                return;
            }
            //LastRowNum 是当前表的总行数-1（注意）
            for (i = 0; i < column_num; i++)
            {
                value_tmp = row.GetCell(i).ToString();
                if (string.Compare(value_tmp, info_names[i]) != 0)
                {
                    file_ok = false;
                    break;
                }
                //value_total += value_tmp + "\r\n";
            }
            if (!file_ok)
            {
                MessageBox.Show("第" + i + "列名称错误！请修改后尝试");
                return;
            }
            /*判断文件名是否有问题*/
            out_file_path = outputfile_foilder + "\\" +textBox2.Text.ToString()+".docx";//注意组合后的路径名称
            if(File.Exists(out_file_path.ToString()))
            {
                MessageBox.Show("输入的文件已经存在，请重新输入！");
                return;
            }
            //先定义一个外面的文件
            /*循环遍历每一行，写入一个文件即可*/
            newWordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);//初始化输出文件
            /*循环体*/
            for (i=1;i< line_num ;i++)  
            {
                label2.Text = "正在处理第"+i+"条信息……";
                /*变量区*/
                //string name_now, workplace_now, retire_year, retire_money_time, submit_time, declare_time, last_sign_time;
                row = sheet.GetRow(i);
                for(j=0;j<column_num;j++)
                {
                    info_tmp[j] = row.GetCell(j).ToString();
                    textBox1.Text += "\r\n" + info_tmp[j];
                }
                use_str_info();//处理字符串到需要的格式
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
            }
            //循环结束后保存以及关闭文件
            //WdSaveFormat为Word 默认文档的保存格式
            object format = MSWord.WdSaveFormat.wdFormatDocumentDefault;// office 2007就是wdFormatDocumentDefault
            newWordDoc.SaveAs(ref out_file_path, ref format, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            newWordDoc.Close();
            label2.Text = "处理完成";
        }
    }
}
