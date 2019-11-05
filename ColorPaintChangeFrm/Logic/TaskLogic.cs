using System.Data;
using System.Threading;

namespace ColorPaintChangeFrm.Logic
{
    public class TaskLogic
    {
        ImportDt importDt = new ImportDt();
        GenerateDt generateDt = new GenerateDt();
        ExportDt exportDt = new ExportDt();

        #region 变量参数
        private int _taskid;
        private string _fileAddress;       //文件地址
        private DataTable _dt;             //获取dt(从EXCEL获取的DT)
        private DataTable _tempdt;         //保存运算成功的DT(导出时使用)
        private int _seletcomid;           //获取下拉框所选的值ID(导出时使用)

        private DataTable _resultTable;   //返回DT
        private bool _resultMark;        //返回是否成功标记
        #endregion

        #region Set
        /// <summary>
        /// 中转ID
        /// </summary>
        public int TaskId { set { _taskid = value; } }

        /// <summary>
        /// //接收文件地址信息
        /// </summary>
        public string FileAddress { set { _fileAddress = value; } }

        /// <summary>
        /// 获取dt(从EXCEL获取的DT)
        /// </summary>
        public DataTable Data { set { _dt = value; } }
        /// <summary>
        /// 获取下拉框所选的值ID(导出时使用)
        /// </summary>
        public int Selectcomid { set { _seletcomid = value; } }
        #endregion

        #region Get
        /// <summary>
        ///返回DataTable至主窗体
        /// </summary>
        public DataTable RestulTable => _resultTable;

        /// <summary>
        ///  返回是否成功标记
        /// </summary>
        public bool ResultMark => _resultMark;
        /// <summary>
        /// 返回运算成功的表头DT(导出时使用)
        /// </summary>
        public DataTable Tempdt => _tempdt;
        #endregion

        public void StartTask()
        {
            Thread.Sleep(1000);

            switch (_taskid)
            {
                //导入
                case 0:
                    OpenExcelImporttoDt(_fileAddress, _seletcomid);
                    break;
                //运算
                case 1:
                    GenerateRecord(_seletcomid, _dt);
                    break;
                //导出
                case 2:
                    ExportDtToExcel(_fileAddress, _tempdt, _seletcomid);
                    break;
            }
        }

        /// <summary>
        /// 导入
        /// </summary>
        /// <param name="fileAddress"></param>
        /// <param name="selectid">获取下拉框所选的值ID 1:导出至旧数据库 2:导出至新数据库 3:以横向方式导出至新数据库模板</param>
        private void OpenExcelImporttoDt(string fileAddress, int selectid)
        {
            _resultTable = importDt.OpenExcelImporttoDt(fileAddress, selectid);
        }

        /// <summary>
        /// 运算
        /// </summary>
        /// <param name="selectid">获取下拉框所选的值ID 1:导出至旧数据库 2:导出至新数据库 3:以横向方式导出至新数据库模板</param>
        /// <param name="dt">从EXCEL导入过来的DT</param>
        private void GenerateRecord(int selectid, DataTable dt)
        {
            _tempdt = generateDt.GenerateExcelSourceDt(selectid, dt);
            _resultMark = _tempdt.Rows.Count > 0;
        }

        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="fileAddress"></param>
        /// <param name="tempdt">整理后的DT</param>
        /// <param name="selectid">获取下拉框所选的值ID(导出时使用)</param>
        private void ExportDtToExcel(string fileAddress, DataTable tempdt,int selectid)
        {
            _resultMark = exportDt.ExportDtToExcel(fileAddress,tempdt,selectid);
        }

    }
}
