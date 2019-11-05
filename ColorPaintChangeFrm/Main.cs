﻿using System;
using System.Data;
using System.Threading;
using System.Windows.Forms;
using ColorPaintChangeFrm.Logic;
using Mergedt;

namespace ColorPaintChangeFrm
{
    public partial class Main : Form
    {
        TaskLogic task = new TaskLogic();
        Load load = new Load();

        #region
        //保存EXCEL导入的DT
        private DataTable _importdt;
        //保存运算成功的表头DT(导出时使用)
        private DataTable _tempdt;
        //保存运算成功的表体DT(导出时使用)
        private DataTable _tempdtldt;
        #endregion

        public Main()
        {
            InitializeComponent();
            OnRegisterEvents();
            OnShowTypeList();
        }

        private void OnRegisterEvents()
        {
            btnopen.Click += Btnopen_Click;
            tmclose.Click += Tmclose_Click;
        }

        /// <summary>
        /// 导入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btnopen_Click(object sender, EventArgs e)
        {
            try
            {
                //获取下拉列表信息
                var dvCustidlist = (DataRowView)comselect.Items[comselect.SelectedIndex];
                var selectid = Convert.ToInt32(dvCustidlist["Id"]);

                var openFileDialog = new OpenFileDialog { Filter = "Xlsx文件|*.xlsx" };
                if (openFileDialog.ShowDialog() != DialogResult.OK) return;
                var fileAdd = openFileDialog.FileName;

                //将所需的值赋到Task类内
                task.TaskId = 0;
                task.FileAddress = fileAdd;
                task.Selectcomid = selectid;

                //使用子线程工作(作用:通过调用子线程进行控制Load窗体的关闭情况)
                new Thread(Start).Start();
                load.StartPosition = FormStartPosition.CenterScreen;
                load.ShowDialog();

                _importdt = task.RestulTable;

                if (_importdt.Rows.Count == 0) throw new Exception("不能成功导入EXCEL内容,请检查模板是否正确.");
                else
                {
                    var clickMessage = $"导入成功,是否进行运算功能?";
                    var clickMes = $"运算成功,是否进行导出至Excel?";

                    if (MessageBox.Show(clickMessage, "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    {
                        if (!Generatedt(_importdt)) throw new Exception("运算不成功,请联系管理员");
                        else if (MessageBox.Show(clickMes, "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                        {
                            Exportdt();
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 运算功能
        /// </summary>
        bool Generatedt(DataTable dt)
        {
            var result = true;
            try
            {
                //获取下拉列表信息
                var dvCustidlist = (DataRowView)comselect.Items[comselect.SelectedIndex];
                var selectid = Convert.ToInt32(dvCustidlist["Id"]);

                task.TaskId = 1;
                task.Data = dt;
                task.Selectcomid = selectid;

                //使用子线程工作(作用:通过调用子线程进行控制Load窗体的关闭情况)
                new Thread(Start).Start();
                load.StartPosition = FormStartPosition.CenterScreen;
                load.ShowDialog();

                result = task.ResultMark;
                _tempdt = task.Tempdt;
            }
            catch (Exception)
            {
                result = false;
            }
            return result;
        }


        /// <summary>
        /// 导出功能
        /// </summary>
        void Exportdt()
        {
            try
            {
                //获取下拉列表信息
                var dvCustidlist = (DataRowView)comselect.Items[comselect.SelectedIndex];
                var selectid = Convert.ToInt32(dvCustidlist["Id"]);

                var saveFileDialog = new SaveFileDialog { Filter = "Xlsx文件|*.xlsx" };
                if (saveFileDialog.ShowDialog() != DialogResult.OK) return;
                var fileAdd = saveFileDialog.FileName;

                task.TaskId = 2;
                task.FileAddress = fileAdd;
                task.Selectcomid = selectid;

                //使用子线程工作(作用:通过调用子线程进行控制Load窗体的关闭情况)
                new Thread(Start).Start();
                load.StartPosition = FormStartPosition.CenterScreen;
                load.ShowDialog();

                if (!task.ResultMark) throw new Exception("导出异常");
                else
                {
                    MessageBox.Show($"导出成功!可从EXCEL中查阅导出效果", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 关闭
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Tmclose_Click(object sender, EventArgs e)
        {
            try
            {
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        ///子线程使用(重:用于监视功能调用情况,当完成时进行关闭LoadForm)
        /// </summary>
        private void Start()
        {
            task.StartTask();

            //当完成后将Form2子窗体关闭
            this.Invoke((ThreadStart)(() => {
                load.Close();
            }));
        }

        /// <summary>
        /// 下拉列表初始化
        /// </summary>
        private void OnShowTypeList()
        {
            var dt = new DataTable();

            //创建表头
            for (var i = 0; i < 2; i++)
            {
                var dc = new DataColumn();
                switch (i)
                {
                    case 0:
                        dc.ColumnName = "Id";
                        break;
                    case 1:
                        dc.ColumnName = "Name";
                        break;
                }
                dt.Columns.Add(dc);
            }

            //创建行内容
            for (var j = 0; j < 2; j++)
            {
                var dr = dt.NewRow();

                switch (j)
                {
                    case 0:
                        dr[0] = "1";
                        dr[1] = "以纵向方式导出";
                        break;
                    case 1:
                        dr[0] = "2";
                        dr[1] = "以横向方式导出";
                        break;
                }
                dt.Rows.Add(dr);
            }

            comselect.DataSource = dt;
            comselect.DisplayMember = "Name"; //设置显示值
            comselect.ValueMember = "Id";    //设置默认值内码
        }

    }
}