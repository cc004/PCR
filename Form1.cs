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
        private readonly Dictionary<string, string> unitNames;
        private readonly List<string> UBNames;

        public Form1()
        {
            InitializeComponent();
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            try
            {
                
                if (File.Exists(System.Windows.Forms.Application.StartupPath + "/UnitNameDic.json"))
                {
                    unitNames = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(System.Windows.Forms.Application.StartupPath + "/UnitNameDic.json"));
                }
                else
                {
                    unitNames = JsonConvert.DeserializeObject<Dictionary<string, string>>(Properties.Resources.UnitNameDic);
                }
                if (File.Exists(System.Windows.Forms.Application.StartupPath + "/UBNameDic.json"))
                {
                    UBNames = JsonConvert.DeserializeObject<List<string>>(File.ReadAllText(System.Windows.Forms.Application.StartupPath + "/UBNameDic.json"));
                }
                else
                {
                    UBNames = JsonConvert.DeserializeObject<List<string>>(Properties.Resources.UBNameDic);
                }
            }  
            catch (Exception ex)
            {
                unitNames = JsonConvert.DeserializeObject<Dictionary<string, string>>(Properties.Resources.UnitNameDic);
                UBNames = JsonConvert.DeserializeObject<List<string>>(Properties.Resources.UBNameDic);
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
            if(comboBox2.Text == "m:ss (0:58)")
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
            else if(comboBox2.Text == "mss  (058)")
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
            if(openFileDialog1.ShowDialog() == DialogResult.Cancel)
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
                MessageBox.Show("文件读取失败! 错误信息: " + ex.Message,"错误" , MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            progressBar1.Value = 0;
            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }


            string templateFilePath = System.Windows.Forms.Application.StartupPath + "/轴模板.xlsx";
            string outFileName = saveFileDialog1.FileName;
            if (outFileName == "" || 
                (File.Exists(outFileName) && MessageBox.Show("文件已经存在，继续操作将会覆盖文件！", "文件已存在", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.Cancel))
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

            if (File.Exists(templateFilePath) == false)
            {
                templateFilePath = Path.GetTempPath() + "轴模板.xlsl";
                File.WriteAllBytes(templateFilePath, Properties.Resources.轴模板);
            }

            try
            {

                Workbooks wbks = app.Workbooks;
                _Workbook wbk = wbks.Add(textBox1.Text);
                Sheets shs = wbk.Sheets;

                //打开保存文件
                Worksheet wsh2 = wbks.Add(templateFilePath).Sheets.get_Item("轴");

                //轴编号
                wsh2.Cells[4, 3] = textBox4.Text;
                //轴标题
                wsh2.Cells[4, 5] = textBox2.Text;
                //轴作者
                wsh2.Cells[4, 12] = textBox3.Text;


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
                        if (unitNames.Keys.Contains(text))
                        {
                            wsh2.Cells[6, 5 + (4 - i) * 2] = unitNames[text];
                            units.Add(unitNames[text]);
                        }
                        else
                        {
                            wsh2.Cells[6, 5 + (4 - i) * 2] = text;
                            units.Add(text);
                        }
                        //等级
                        wsh2.Cells[12, 6 + (4 - i) * 2] = wsh.Cells[i + 3, 3].Value;
                        //星级
                        wsh2.Cells[13, 6 + (4 - i) * 2] = wsh.Cells[i + 3, 4].Value;
                        //专武等级
                        wsh2.Cells[15, 6 + (4 - i) * 2] = wsh.Cells[i + 3, 17].Value == 0 ? "-" : wsh.Cells[i + 3, 17].Value;
                        //角色Rank
                        int num = 6;
                        for (int j = 0; j < 6; j++)
                        {
                            if (wsh.Cells[i + 3, 7 + j].Value.ToString() == "未装备")
                            {
                                num--;
                            }
                        }
                        wsh2.Cells[14, 6 + (4 - i) * 2] = wsh.Cells[i + 3, 6].Value.ToString() + "-" + num.ToString();
                    }
                    else
                    {
                        break;
                    }
                }

                //BOSS信息
                text = wsh.Cells[10, 2].Value.ToString();
                unitIds.Insert(0, (int)wsh.Cells[10, 1].Value);
                if (unitNames.Keys.Contains(text))
                {
                    wsh2.Cells[6, 3] = unitNames[text];
                    units.Insert(0, unitNames[text]);
                }
                else
                {
                    wsh2.Cells[6, 3] = text;
                    units.Insert(0, text);
                }


                //设置角色头像
                if (checkBox1.Checked)
                {
                    for (int i = 1; i <= unitIds.Count; i++)
                    {
                        int unitId = unitIds[i - 1];
                        unitId = unitId < 200000 ? unitId + 30 : unitId;
                        string path = System.Windows.Forms.Application.StartupPath + "/images/icon_unit_" + unitId.ToString() + ".png";
                        if (File.Exists(path))
                        {
                            wsh2.Shapes.AddPicture(path, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, (float)(120.0 + (112.5 * (i == 1 ? 0 : 7 - i))), 101, 96, 96);
                        }
                    }
                }

                //读取技能循环
                wsh = (_Worksheet)shs.get_Item("技能循环");

                int rowId = 19;

                for (int i = 1; i < 1000; i++)
                {
                    cell = wsh.Cells[i, 1];

                    if (cell.Value == null || cell.Value.ToString() != "角色技能循环详情")
                    {
                        continue;
                    }
                    for (int j = i + 2; j < 9999999; j++)
                    {
                        cell = wsh.Cells[j, 1];
                        progressBar1.Value = (int)(cell.value / 5400 * 100);
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
                                    wsh2.Cells[rowId, 2] = frame;
                                    wsh2.Cells[rowId, 3] = GetTimeText(frame);
                                    wsh2.Cells[rowId, 5] = units[k];
                                    wsh2.Cells[rowId, 8] = 0;
                                    wsh2.Range["H" + rowId.ToString()].Font.Color = Color.FromArgb(0, 0, 0);

                                    //获取ub伤害
                                   string str = "";
                                    for (int l = j - 1; (int)wsh.Cells[l, 1].Value == frame; l--)
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

                                    for (int l = j + 1; (int)wsh.Cells[l, 1].Value == frame; l++)
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

                                    foreach (Match mat in Regex.Matches(str, @"对目标造成(\d+)点(暴击)?伤害"))
                                    {
                                        int damage = int.Parse(mat.Groups[1].Value);
                                        if (comboBox1.Text == "最高伤害" && wsh2.Cells[rowId, 8].Value < damage)
                                        {
                                            wsh2.Cells[rowId, 8].Value = damage;
                                            wsh2.Range["H" + rowId.ToString()].Font.Color = wsh2.Cells[rowId, 4] = mat.Groups[2].Value == "暴击" ? Color.FromArgb(255, 0, 0) : Color.FromArgb(0, 0, 0);
                                        }else if (comboBox1.Text == "UB总伤害")
                                        {
                                            wsh2.Cells[rowId, 8].Value = (int)wsh2.Cells[rowId, 8].Value + damage;
                                            if(mat.Groups[2].Value == "暴击")
                                            {
                                                wsh2.Range["H" + rowId.ToString()].Font.Color = wsh2.Cells[rowId, 4] = Color.FromArgb(255, 0, 0);
                                            }
                                        }
                                        if (comboBox1.Text == "最低伤害" && (wsh2.Cells[rowId, 8].Value > damage || wsh2.Cells[rowId, 8].Value == 0))
                                        {
                                            wsh2.Cells[rowId, 8].Value = damage;
                                            wsh2.Range["H" + rowId.ToString()].Font.Color = wsh2.Cells[rowId, 4] = mat.Groups[2].Value == "暴击" ? Color.FromArgb(255, 0, 0) : Color.FromArgb(0, 0, 0);
                                        }
                                    }
                                }
                                else
                                {
                                    wsh2.Range["B" + rowId.ToString() + ":" + "N" + rowId.ToString()].Clear();
                                    wsh2.Range["B99:N99"].Copy(wsh2.Range["B" + rowId.ToString()]);
                                    wsh2.Cells[rowId, 2] = frame;
                                    wsh2.Cells[rowId, 3] = GetTimeText(frame); ;
                                }
                                rowId++;
                                break;
                            }
                        }
                    }
                    wsh2.Range["B" + rowId.ToString() + ":" + "N" + rowId.ToString()].Clear();
                    wsh2.Range["B100:N100"].Copy(wsh2.Range["B" + rowId.ToString()]);
                    wsh2.Cells[rowId, 2].Value = 5399;
                    wsh2.Cells[rowId, 3].Value = GetTimeText(5399);
                    rowId++;
                    wsh2.Range["B" + rowId.ToString() + ":" + "N100"].Clear();
                    wsh2.Range["C4"].Select();
                    break;
                }


                app.DisplayAlerts = false;
                wsh2.SaveAs(outFileName);
                app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                if (MessageBox.Show("是否打开文件", "生成成功", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.OK)
                {
                    Process.Start(outFileName);
                }
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show("错误信息: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            app.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
        }
    }
}
