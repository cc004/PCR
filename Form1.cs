using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace 写轴器
{
    public partial class Form1 : Form
    {
        private Dictionary<string, string> unitNames;
        private List<string> UBNames;
        private Dictionary<string, string> setting;
        public void LoadTemplate(string template)
        {
            Application app = null;
            try
            {
                app = app = new Application();
                unitNames = new Dictionary<string, string> { };
                UBNames = new List<string> { };
                setting = new Dictionary<string, string> { };
                _Workbook wbk = app.Workbooks.Add(template);
                Sheets shs = wbk.Sheets;
                _Worksheet wsh = (_Worksheet)shs.get_Item("设置");


                for (int i = 2; ; i++)
                {
                    var cell = wsh.Cells[i, 5];
                    if (cell.value != null)
                    {
                        UBNames.Add(cell.value.ToString());
                        cell = wsh.Cells[i, 3];
                        if (cell.value != null)
                        {
                            unitNames.Add(cell.value.ToString(), wsh.Cells[i, 4].value.ToString());
                        }
                    }
                    else
                    {
                        break;
                    }
                }
                for (int i = 2; ; i++)
                {
                    var cell = wsh.Cells[i, 6];
                    if (cell.value != null)
                    {
                        unitNames.Add(cell.value.ToString(), wsh.Cells[i, 7].value.ToString());
                    }
                    else
                    {
                        break;
                    }
                }
                for (int i = 2; ; i++)
                {
                    var cell = wsh.Cells[i, 1];
                    if (cell.value != null)
                    {
                        setting.Add(cell.value.ToString(), wsh.Cells[i, 2].value == null ? null : wsh.Cells[i, 2].value.ToString());
                    }
                    else
                    {
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                throw ex;
            }
            finally
            {
                app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            }
        }

        public Form1()
        {
            InitializeComponent();
            unitNames = new Dictionary<string, string> { };
            UBNames = new List<string> { };
            setting = new Dictionary<string, string> { };

            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;

            try
            {
                foreach (string file in Directory.GetFiles(System.Windows.Forms.Application.StartupPath + "/轴模板/"))
                {
                    string[] str = file.Split('/', '\\');
                    comboBox3.Items.Add(str[str.Count() - 1]);
                }
                comboBox3.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("初始化失败! 错误: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }


        public string GetTimeText(int frame)
        {
            //m:ss (0:58)
            //m-ss (0-58)
            //mss  (058)
            //mmss (0058)
            //mm:ss(00:58)
            //mm-ss(00-58)
            //帧   (1860)
            int time = (int)Math.Ceiling(((5400.0 - frame) / 60));
            string text = frame.ToString("0000");
            if (comboBox2.Text == "m:ss (0:58)")
            {
                text = (time / 60).ToString() + ":" + (time % 60).ToString("00");
            }
            else if (comboBox2.Text == "m-ss (0-58)")
            {
                text = (time / 60).ToString() + "-" + (time % 60).ToString("00");
            }
            else if (comboBox2.Text == "m-ss (0-58)")
            {
                text = (time / 60).ToString() + "-" + (time % 60).ToString("00");
            }
            else if (comboBox2.Text == "mss  (058)")
            {
                text = (time / 60).ToString() + (time % 60).ToString("00");
            }
            else if (comboBox2.Text == "mmss (0058)")
            {
                text = (time / 60).ToString("00") + (time % 60).ToString("00");
            }
            else if (comboBox2.Text == "mm:ss(00:58)")
            {
                text = (time / 60).ToString("00") + ":" + (time % 60).ToString("00");
            }
            else if (comboBox2.Text == "mm-ss(00-58)")
            {
                text = (time / 60).ToString("00") + "-" + (time % 60).ToString("00");
            }

            return text;
        }




        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }
            string fileName = openFileDialog1.FileName;
            if (fileName == "")
            {
                return;
            }
            Application app = null;
            try
            {
                app = new Application();
            }
            catch (Exception ex)
            {
                MessageBox.Show("初始化失败! 可能是未安装Microsoft.Office.\n详细错误信息: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                Workbooks wbks = app.Workbooks;
                _Workbook wbk = wbks.Add(fileName);
                Sheets shs = wbk.Sheets;
                _Worksheet wsh = (_Worksheet)shs.get_Item("基础数据");
                var cell = wsh.Cells[1, 1];
                if (cell != null && cell.Value != null && cell.Value.ToString() == "角色基础参数")
                {
                    string text = "";
                    for (int i = 0; i < 5; i++)
                    {
                        cell = wsh.Cells[i + 3, 2];
                        if (cell != null && cell.Value != null)
                        {
                            string name = cell.Value.ToString();
                            if (unitNames.Keys.Contains(name))
                            {
                                text += unitNames[name];
                            }
                            else
                            {
                                text += name;
                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                    textBox1.Text = fileName;
                    textBox2.Text = text;
                    var mat = Regex.Match(fileName, @"\\([A-Ea-e][1-6])\-");
                    textBox4.Text = mat.Success ? mat.Groups[1].Value : "轴编号";
                    button2.Enabled = true;
                }
                else
                {
                    throw new Exception("未找到角色基础参数.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("文件读取失败! 错误信息: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button2.Enabled = false;
            }
            app.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
        }


        private void button2_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            start();
            this.Enabled = true;
        }


        public void start()
        {
            List<string> units = new List<string>();
            List<int> unitIds = new List<int>();
            List<string> _units = new List<string>();
            progressBar1.Value = 0;

            string templateFilePath = System.Windows.Forms.Application.StartupPath + "/轴模板/" + comboBox3.Text;

            string outFilePath = textBox1.Text;

            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }

            outFilePath = saveFileDialog1.FileName;
            if (outFilePath == "" ||
                (File.Exists(outFilePath) && MessageBox.Show("文件已经存在，继续操作将会覆盖文件！", "文件已存在", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.Cancel))
            {
                return;
            }



            System.Windows.Forms.Application.DoEvents();

            Application app = null;
            try
            {
                app = new Application();
            }
            catch (Exception ex)
            {
                MessageBox.Show("初始化失败! 可能是未安装Microsoft.Office.\n详细错误信息: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }



            try
            {
                Workbooks wbks = app.Workbooks;
                _Workbook wbk = wbks.Add(textBox1.Text);
                Sheets shs = wbk.Sheets;

                //打开保存文件
                _Workbook wbk2 = wbks.Add(templateFilePath);
                Worksheet wsh2 = wbk2.Sheets.get_Item("轴");


                //轴编号
                if (setting["轴编号坐标"] != null)
                {
                    wsh2.Range[setting["轴编号坐标"]].Value = textBox4.Text;
                }

                //轴标题
                if (setting["轴标题坐标"] != null)
                {
                    wsh2.Range[setting["轴标题坐标"]].Value = textBox2.Text;
                }
                //轴作者
                if (setting["轴作者坐标"] != null)
                {
                    wsh2.Range[setting["轴作者坐标"]].Value = textBox3.Text;
                }

                //角色信息
                _Worksheet wsh = (_Worksheet)shs.get_Item("基础数据");
                var cell = wsh.Cells[1, 1];
                string text = "";
                for (int i = 0; i < 5; i++)
                {
                    cell = wsh.Cells[i + 3, 2];
                    if (cell != null && cell.Value != null)
                    {
                        text = cell.Value.ToString();
                        unitIds.Add((int)wsh.Cells[i + 3, 1].Value);
                        _units.Add(text);
                        if (unitNames.Keys.Contains(text))
                        {
                            units.Add(unitNames[text]);
                            text = unitNames[text];
                        }
                        else
                        {
                            units.Add(text);
                        }
                        if (setting["角色" + (i + 1) + "名称坐标"] != null)
                        {
                            wsh2.Range[setting["角色" + (i + 1) + "名称坐标"]].Value = text;
                        }

                        //等级
                        string str = setting["角色" + (i + 1) + "等级坐标"];
                        if (str != null)
                        {
                            text = wsh2.Range[str].Value == null ? "" : wsh2.Range[str].Value.ToString() + " ";
                            wsh2.Range[str].Value = text + String.Format(setting["等级文本"], wsh.Cells[i + 3, 3].Value);
                        }

                        //星级
                        str = setting["角色" + (i + 1) + "星级坐标"];
                        if (str != null)
                        {
                            text = wsh2.Range[str].Value == null ? "" : wsh2.Range[str].Value.ToString() + " ";
                            wsh2.Range[str].Value = text + String.Format(setting["星级文本"], wsh.Cells[i + 3, 4].Value);
                        }


                        //角色Rank
                        int num = 6;
                        for (int j = 0; j < 6; j++)
                        {
                            if (wsh.Cells[i + 3, 7 + j].Value.ToString() == "未装备")
                            {
                                num--;
                            }
                        }

                        str = setting["角色" + (i + 1) + "Rank坐标"];
                        if (str != null)
                        {
                            text = wsh2.Range[str].Value == null ? "" : wsh2.Range[str].Value.ToString() + " ";
                            wsh2.Range[str].Value = text + String.Format(setting["Rank文本"], wsh.Cells[i + 3, 6].Value.ToString() + "-" + num.ToString());

                        }

                        //专武等级
                        str = setting["角色" + (i + 1) + "专武坐标"];
                        if (str != null)
                        {

                            text = wsh2.Range[str].Value == null ? "" : wsh2.Range[str].Value.ToString() + " ";
                            if (wsh.Cells[i + 3, 17].Value == 0)
                            {
                                if (setting["无专武文本"] != null)
                                {
                                    wsh2.Range[str].Value = text + setting["无专武文本"];
                                }


                            }
                            else
                            {
                                wsh2.Range[str].Value = text + String.Format(setting["专武文本"], wsh.Cells[i + 3, 17].Value);
                            }

                        }

                    }
                    else
                    {
                        break;
                    }
                }

                //BOSS信息
                string bossName = wsh.Cells[10, 2].Value.ToString();
                unitIds.Insert(0, (int)wsh.Cells[10, 1].Value);
                if (unitNames.Keys.Contains(bossName))
                {
                    text = unitNames[bossName];
                    units.Insert(0, unitNames[bossName]);
                }
                else
                {
                    text = bossName;
                    units.Insert(0, bossName);
                    _units.Insert(0, bossName);
                }
                if (setting["BOSS名称坐标"] != null)
                {
                    wsh2.Range[setting["BOSS名称坐标"]].Value = text;
                }


                //设置角色头像
                if (checkBox1.Checked)
                {
                    for (int i = 0; i < unitIds.Count; i++)
                    {
                        int unitId = unitIds[i];
                        unitId = unitId < 200000 ? unitId + 30 : unitId;
                        string path = System.Windows.Forms.Application.StartupPath + "/images/icon_unit_" + unitId.ToString() + ".png";
                        if (File.Exists(path))
                        {
                            if (i == 0 && setting["BOSS头像"] != null)
                            {
                                string[] arr = setting["BOSS头像"].Split(',');
                                wsh2.Shapes.AddPicture(path, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, float.Parse(arr[0]), float.Parse(arr[1]), float.Parse(arr[2]), float.Parse(arr[3]));
                            }
                            else if (i != 0 && setting["角色" + i + "头像"] != null)
                            {
                                string[] arr = setting["角色" + i + "头像"].Split(',');
                                wsh2.Shapes.AddPicture(path, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, float.Parse(arr[0]), float.Parse(arr[1]), float.Parse(arr[2]), float.Parse(arr[3]));
                            }
                        }
                    }
                }

                //读取技能循环
                wsh = (_Worksheet)shs.get_Item("技能循环");

                int rowId = int.Parse(setting["轴起始行"]);

                for (int i = 1; i < 1000; i++)
                {
                    cell = wsh.Cells[i, 1];

                    if (cell.Value == null || cell.Value.ToString() != "角色技能循环详情")
                    {
                        continue;
                    }

                    string oldTimeText = "";

                    for (int j = i + 2; j < 9999999; j++)
                    {
                        cell = wsh.Cells[j, 1];


                        if ((int)(cell.value / 5400 * 100) != progressBar1.Value)
                        {
                            progressBar1.Value = (int)(cell.value / 5400 * 100);
                        }

                        System.Windows.Forms.Application.DoEvents();
                        if (cell.value == 5400)
                        {
                            break;
                        }
                        int frame = (int)cell.Value;
                        for (int k = 0; k < 6; k++)
                        {
                            cell = wsh.Cells[j, 2 + k * 5];
                            if (cell.Value != null && cell.Value.ToString() == "切换到UB状态")
                            {
                                Debug.WriteLine("[" + frame.ToString() + "]" + units[k] + "释放UB");
                                if (k != 0)
                                {
                                    if (setting["帧数列"] != null)
                                    {
                                        wsh2.Cells[rowId, int.Parse(setting["帧数列"])] = frame;
                                    }
                                    string timeText = GetTimeText(frame);
                                    wsh2.Cells[rowId, int.Parse(setting["角色列"])] = units[k];
                                    wsh2.Cells[rowId, int.Parse(setting["伤害列"])] = 0;



                                    if (oldTimeText != timeText)
                                    {
                                        wsh2.Cells[rowId, int.Parse(setting["时间列"])] = timeText;
                                        oldTimeText = timeText;
                                    }
                                    else
                                    {
                                        wsh2.Cells[rowId, int.Parse(setting["时间列"])] = null;
                                    }


                                    //获取ub伤害
                                    string str = "";
                                    int pt = 0;
                                    for (int l = j - 1; (int)wsh.Cells[l, 1].Value == frame; l--)
                                    {
                                        pt = l;
                                    }

                                    for (int l = (pt != 0 ? pt : j); (int)wsh.Cells[l, 1].Value == frame; l++)
                                    {
                                        cell = wsh.Cells[l, 2 + k * 5];
                                        if (cell.Value != null)
                                        {
                                            text = cell.Value.ToString().Replace("释放技能", "");
                                            if (UBNames.Contains(text))
                                            {
                                                str += wsh.Cells[l, 4 + k * 5].Value.ToString() + "; ";
                                            }
                                        }
                                    }


                                    Debug.WriteLine(str);
                                    string tmp = "";
                                    int id = 0;
                                    foreach (Match mat in Regex.Matches(str, @"(目标：([^\,]+),|/)对目标造成(\d+)点(暴击)?伤害"))
                                    {
                                        if (mat.Groups[1].Value.ToString() != "/")
                                        {
                                            tmp = mat.Groups[2].Value.ToString();
                                        }


                                        if (tmp != _units[k])
                                        {
                                            int damage = int.Parse(mat.Groups[3].Value);
                                            string text2 = mat.Groups[4].Value;

                                            int unitId = unitIds[k];

                                            if ((unitId == 101601 || unitId == 107101 || (unitId == 170101 && comboBox1.Text != "最低伤害")) == false && text2 == "暴击")
                                            {
                                                text2 = "";
                                                damage /= 2;
                                            }


                                            if (comboBox1.Text == "最高伤害" && wsh2.Cells[rowId, int.Parse(setting["伤害列"])].Value < damage)
                                            {
                                                wsh2.Cells[rowId, int.Parse(setting["伤害列"])].Value = damage;
                                                //wsh2.Range["H" + rowId.ToString()].Font.Color = wsh2.Cells[rowId, 4] = text2 == "暴击" ? Color.FromArgb(255, 0, 0) : Color.FromArgb(0, 0, 0);
                                            }
                                            else if (comboBox1.Text == "UB总伤害")
                                            {
                                                wsh2.Cells[rowId, int.Parse(setting["伤害列"])].Value = (int)wsh2.Cells[rowId, int.Parse(setting["伤害列"])].Value + damage;
                                                if (text2 == "暴击")
                                                {
                                                    //wsh2.Range["H" + rowId.ToString()].Font.Color = wsh2.Cells[rowId, 4] = Color.FromArgb(255, 0, 0);
                                                }
                                            }
                                            if (comboBox1.Text == "最低伤害" && (wsh2.Cells[rowId, int.Parse(setting["伤害列"])].Value > damage || wsh2.Cells[rowId, int.Parse(setting["伤害列"])].Value == 0))
                                            {
                                                wsh2.Cells[rowId, int.Parse(setting["伤害列"])].Value = damage;
                                                //wsh2.Range["H" + rowId.ToString()].Font.Color = wsh2.Cells[rowId, 4] = text2 == "暴击" ? Color.FromArgb(255, 0, 0) : Color.FromArgb(0, 0, 0);
                                            }
                                        }
                                        id++;
                                    }
                                    if (wsh2.Cells[rowId, int.Parse(setting["伤害列"])].Value == 0)
                                    {
                                        wsh2.Cells[rowId, int.Parse(setting["伤害列"])].Value = null;
                                    }
                                }
                                else
                                {
                                    wsh2.Range["A" + rowId.ToString() + ":" + "AAA" + rowId.ToString()].Clear();
                                    wsh2.Range["A" + setting["BOOSUB行"] + ":" + "AAA" + setting["BOOSUB行"]].Copy(wsh2.Range["A" + rowId.ToString() + ":" + "AAA" + rowId.ToString()]);
                                    if (setting["帧数列"] != null)
                                    {
                                        wsh2.Cells[rowId, int.Parse(setting["帧数列"])] = frame;
                                    }

                                    string timeText = GetTimeText(frame);

                                    if (wsh2.Cells[rowId, int.Parse(setting["时间列"])].value == null)
                                    {
                                        wsh2.Cells[rowId, int.Parse(setting["时间列"])] = timeText;
                                    }
                                    else
                                    {
                                        wsh2.Cells[rowId, int.Parse(setting["时间列"])] = timeText + " " + wsh2.Cells[rowId, int.Parse(setting["时间列"])].value.ToString();
                                    }

                                    oldTimeText = timeText;
                                }
                                rowId++;
                                break;
                            }
                        }
                    }
                    //wsh2.Range["A" + rowId.ToString() + ":" + "AAA" + rowId.ToString()].Clear();
                    wsh2.Range["A" + (int.Parse(setting["轴结尾行"]) - 2) + ":AAA" + setting["轴结尾行"]].Copy(wsh2.Range["A" + rowId.ToString()]);
                    //wsh2.Cells[rowId, int.Parse(setting["帧数列"])].Value = 5400;
                    //wsh2.Cells[rowId, 3].Value = GetTimeText(5400);
                    rowId += 3;
                    wsh2.Range["A" + rowId.ToString() + ":" + "AAA999"].Clear();
                    //wsh2.Range["B2"].Select();
                    break;
                }



                wbk2.Worksheets["设置"].Delete();
                app.DisplayAlerts = false;
                wbk2.SaveAs(outFilePath);
                app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                this.Enabled = true;
                if (MessageBox.Show("是否打开文件", "生成成功", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.OK)
                {
                    Process.Start(outFilePath);
                }
                return;
            }
            catch (Exception ex)
            {
                this.Enabled = true;
                MessageBox.Show("错误信息: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            app.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

            try
            {
                LoadTemplate(System.Windows.Forms.Application.StartupPath + "/轴模板/" + comboBox3.SelectedItem.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show("错误信息: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
