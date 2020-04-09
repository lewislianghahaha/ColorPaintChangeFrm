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
        private int _typeid;               //获取导入的来源(1:常规纵向EXCEL导入 2:带空白纵向EXCEL导入)
        private int _genid;                //获取转换单位(1:按KG进行计算色母量 2:按L进行计算色母量 0:不需计算色母量)
        private int _sortid;               //获取选择条件(1:不筛选 2:获取1个色母数的配方 3:获取2个色母数的配方 4:获取3个色母数并包含PC-60的配方 5:获取包含PC-60并占比>=20%的配方)

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
        /// <summary>
        /// 获取导入的来源(1:常规纵向EXCEL导入 2:带空白纵向EXCEL导入)
        /// </summary>
        public int Typeid { set { _typeid = value; } }
        /// <summary>
        /// 获取转换单位(1:按KG进行计算色母量 2:按L进行计算色母量)
        /// </summary>
        public int Genid { set { _genid = value; } }
        /// <summary>
        /// 获取选择条件(1:不筛选 2:获取1个色母数的配方 3:获取2个色母数的配方 4:获取3个色母数并包含PC-60的配方 5:获取包含PC-60并占比>=20%的配方)
        /// </summary>
        public int Sortid { set { _sortid = value; } }
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
                    OpenExcelImporttoDt(_fileAddress, _typeid);
                    break;
                //导入-含空白纵向EXCEL
                case 3:
                    ImportEmportExcelToDt(_fileAddress, _typeid);
                    break;
                //运算
                case 1:
                    GenerateRecord(_genid,_seletcomid, _sortid,_dt);
                    break;
                //导出
                case 2:
                    ExportDtToExcel(_fileAddress, _tempdt, _seletcomid);
                    break;
            }
        }

        /// <summary>
        /// 导入(常规纵向EXCEL导入)
        /// </summary>
        /// <param name="fileAddress"></param>
        /// <param name="typeid">导入类型=>1:常规纵向EXCEL导入 2:带空白纵向EXCEL导入</param>
        private void OpenExcelImporttoDt(string fileAddress, int typeid)
        {
            _resultTable = importDt.OpenExcelImporttoDt(fileAddress, typeid);
        }

        /// <summary>
        /// 导入(带空白纵向EXCEL导入)
        /// </summary>
        /// <param name="fileAddress"></param>
        /// <param name="typeid">导入类型=>1:常规纵向EXCEL导入 2:带空白纵向EXCEL导入</param>
        private void ImportEmportExcelToDt(string fileAddress, int typeid)
        {
            _resultTable = importDt.OpenExcelImporttoDt(fileAddress, typeid);
        }

        /// <summary>
        /// 运算
        /// </summary>
        /// <param name="genid">转换单位=>1:按KG进行计算色母量 2:按L进行计算色母量</param>
        /// <param name="selectid">获取下拉框所选的值ID 1:以纵向方式导出 2:以横向方式导出 0:不需计算色母量</param>
        /// <param name="sortid">获取选择条件(1:不筛选 2:获取1个色母数的配方 3:获取2个色母数的配方 4:获取3个色母数并包含PC-60的配方 5:获取包含PC-60并占比>=20%的配方)</param>
        /// <param name="dt">从EXCEL导入过来的DT</param>
        private void GenerateRecord(int genid,int selectid, int sortid, DataTable dt)
        {
            _tempdt = generateDt.GenerateExcelSourceDt(genid,selectid, sortid,dt);
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
