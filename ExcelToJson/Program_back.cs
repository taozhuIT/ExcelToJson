/*
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Data;
using Newtonsoft.Json;
using System.Data.OleDb;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Security.AccessControl;
    
/// <summary>
/// 尝试另外一种数据格式，DataTable()和DataRow直接序列化成json,而不用List或者Dictionary来直接序列化成json
/// </summary>
namespace ExcelToJson
{

    class Program
    {
        /// <summary>
        /// 程序入口
        /// </summary>
        /// <param name="args"></param>
        private static void Main(string[] args)
        {
            ExcelToJson tojson = new ExcelToJson();
            tojson.Init();
        }
    }

    /// <summary>
    /// Excel转Json工具类
    /// </summary>
    class ExcelToJson
    {
        // Excel文件列表
        List<FileSystemInfo> infoList = new List<FileSystemInfo>();

        /// <summary>
        /// 工具初始化
        /// </summary>
        public void Init()
        {
            OnStartExcelToJson();
        }
        /// <summary>
        /// 开始进行文件转换
        /// </summary>
        private void OnStartExcelToJson()
        {
            // 获取程序运行路径
            string rootPath = Directory.GetCurrentDirectory();

            // 获取所有的.xlsx文件
            infoList.Clear();
            OnGetCurDirectExcle(rootPath);
            Console.Write("\n");

            foreach (FileSystemInfo info in infoList)
            {
                Console.Write("--------转换:  " + info.Name + "  到json文件开始--------\n");
                // 获取单个表文件信息
                // 连接字符串 Office 07及以上版本 不能出现多余的空格 而且分号注意
                string connstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + info.FullName + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"; 

                OleDbConnection oledbCon = new OleDbConnection(connstring);
                oledbCon.Open();

                // 获取Excel文件的表数据
                DataTable tables = oledbCon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                List<DataTable> tableList = new List<DataTable>();

                // ------------------------得到Excel文件的Table/Sheet信息-----------------
                foreach (DataRow tablesRow in tables.Rows)
                {
                    string sheetTableName = tablesRow["TABLE_NAME"].ToString();
                    // 过滤无效SheetName   
                    if (sheetTableName.Contains("$") && sheetTableName.Replace("'", "").EndsWith("$"))
                    {
                        DataTable columns = oledbCon.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, sheetTableName, null });
                        // 工作表列数
                        if (columns.Rows.Count < 2)                     
                            continue;

                        OleDbCommand cmd = new OleDbCommand("select * from [" + sheetTableName + "] where F1 is not null", oledbCon);
                        OleDbDataAdapter apt = new OleDbDataAdapter(cmd);

                        // 创建表的Table信息
                        DataTable tableDt = new DataTable();
                        apt.Fill(tableDt);
                        tableDt.TableName = sheetTableName.Replace("$", "").Replace("'", "");

                        tableList.Add(tableDt);
                    }
                }
                // ------------------------得到Excel文件的Table/Sheet信息-----------------
                Console.Write("获取  " + info.Name + "  Table/Sheet信息完成\n");

                

                // ------------------------创建序列化Json的数据结构------------------------
                // Json数据使用的key
                List<string> keys = new List<string>();
                // 生成Json文件的数据信息结构(整张表的json结构)
                Dictionary<string, DataTable> chartJsonStruct = new Dictionary<string, DataTable>();

                for(int i = 0; i < tableList.Count; ++i)
                {
                    DataTable tableDt = tableList[i];

                    if (tableDt.Rows.Count < 2)
                        throw new Exception("表必须包含数据");

                    // 得到表格式
                    DataTable dataTableStruct = new DataTable();

                    // table中的数据行
                    for (int n = 0; n < tableDt.Rows.Count; ++n)
                    {
                        DataRow headRow = tableDt.Rows[n];

                        if (n > 0)
                        {
                            if(n == 1)
                            {
                                keys.Clear();
                                for (int j = 0; j < headRow.ItemArray.Length; ++j)
                                {
                                    string key = headRow.ItemArray[j].ToString();

                                    if(key != null && key != "")
                                        keys.Add(headRow.ItemArray[j].ToString());
                                }
                            }
                            else
                            {

                                if (!dataTableStruct.Columns.Contains(keys[0]))
                                {
                                    for (int x = 0; x < headRow.ItemArray.Length; ++x)
                                    {
                                        object valObj = headRow.ItemArray[x];

                                        if (valObj.ToString() != null && valObj.ToString() != " ")
                                        {
                                            string key = keys[x];

                                            if (Util.IsNumber(valObj.ToString()))
                                                dataTableStruct.Columns.Add(key, Type.GetType("System.Int32"));
                                            else if (Util.IsBool(valObj.ToString()))
                                                dataTableStruct.Columns.Add(key, Type.GetType("System.Boolean"));
                                            else if (Util.IsFloat(valObj.ToString()))
                                                dataTableStruct.Columns.Add(key, Type.GetType("System.Single"));
                                            else
                                                dataTableStruct.Columns.Add(key, Type.GetType("System.String"));
                                        }
                                    }
                                }

                                OnFillHorizontaData(headRow, keys, dataTableStruct);
                            }
                        }
                    }

                    chartJsonStruct.Add(tableDt.TableName, dataTableStruct);
                }

                Console.Write("创建  " + info.Name + "  Json数据结构完成\n");
                OnJsonSerializer(info, chartJsonStruct);
                Console.Write("--------转换:  " + info.Name + "  到Json文件完成--------\n\n");
            }
        }

        /// <summary>
        /// 得到Excel表中一个页卡的Json结构
        /// </summary>
        private DataTable OnGetTableStruct(DataTable tableDt_, Dictionary<string, DataTable> chartJsonStruct_)
        {
            DataTable tableStruct = new DataTable();

            if (chartJsonStruct_.ContainsKey(tableDt_.TableName))
            {
                tableStruct = chartJsonStruct_[tableDt_.TableName];
            }
            else
            {
                tableStruct = new DataTable();
                chartJsonStruct_.Add(tableDt_.TableName, tableStruct);
            }

            return tableStruct;
        }

        /// <summary>
        /// 填充json一行的数据
        /// </summary>
        private void OnFillHorizontaData(DataRow headRow_, List<string> keys_, DataTable tableJsonStruct_)
        {
            DataRow dr = tableJsonStruct_.NewRow();

            for (int x = 0; x < headRow_.ItemArray.Length; ++x)
            {
                object valObj = headRow_.ItemArray[x];
                string valStr = valObj.ToString();
                if (valStr != null && valStr != " " && valStr != "")
                {
                    string key = keys_[x];

                    if (Util.IsNumber(valObj.ToString()))
                    {
                        dr[key] = int.Parse(headRow_.ItemArray[x].ToString());
                    }
                    else if (Util.IsFloat(valObj.ToString()))
                    {
                        dr[key] = float.Parse(headRow_.ItemArray[x].ToString());
                    }
                    else if (Util.IsBool(valObj.ToString()))
                    {
                        dr[key] = true;
                    }
                    else
                    {

                        dr[key] = headRow_.ItemArray[x].ToString();
                    }
                }
            }

            // 往Table结构中填充一行数据的结构
            tableJsonStruct_.Rows.Add(dr);
        }

        /// <summary>
        /// 序列化数据成json文件
        /// </summary>
        private void OnJsonSerializer(FileSystemInfo fileInfo_, Dictionary<string, DataTable> chartJsonStruct_)
        {
            Console.Write("开始  "+ fileInfo_.Name+ "  文件数据json序列化\n");
            
            try
            {
                // 原始Excel文件路径
                string origPath = fileInfo_.FullName;

                // 存放Json文件的路径
                string stocPath = null;
                string reppath = fileInfo_.FullName.Substring(0, origPath.LastIndexOf(@"\"));
                string repName = fileInfo_.Name.Replace(".xlsx", ".json");
                stocPath = (reppath + "\\Json\\" + repName);

                // 创建存放Json文件的文件夹
                string jsonFolderPath = stocPath.Substring(0, stocPath.LastIndexOf(@"\"));
                if (!Directory.Exists(jsonFolderPath))
                    Directory.CreateDirectory(jsonFolderPath);

                // 创建json文件
                if (File.Exists(stocPath))
                    File.Delete(stocPath);
                FileStream fs = File.Create(stocPath);
                
                // 写入数据
                // 序列化json数据
                string SerializeJson = JsonConvert.SerializeObject(chartJsonStruct_);
                // 格式json字符串（不格式话，显示就是一行显示所有数据）
                string convertJson = ConvertJsonString(SerializeJson);

                // 得到byte
                byte[] byteJson = Encoding.UTF8.GetBytes(convertJson);
                fs.Write(byteJson, 0, byteJson.Length);
                fs.Close();
            }
            catch (Exception except)
            {
                // 文件写入权限问题
                if (except.GetType() == typeof(System.UnauthorizedAccessException))
                {
                    Console.Write("\n无法创建json文件，请右键以管理员身份运行\n");
                    Console.ReadLine();
                }
            }
        }

        /// <summary>
        /// 获取Conf文件夹目录下的所有Excle文件
        /// </summary>
        /// <param name="rootPath_"></param>
        private List<FileSystemInfo> OnGetCurDirectExcle(string rootPath_)
        {
            // 获取程序运行目录下的所有文件和子目录
            DirectoryInfo root = new DirectoryInfo(rootPath_);
            FileSystemInfo[] files = root.GetFileSystemInfos();

            foreach(FileSystemInfo info in files)
            {
                // 判断是否是文件夹
                if (Directory.Exists(info.FullName))
                {
                    OnGetCurDirectExcle(info.FullName);
                }
                else if(info.Extension == ".xlsx")
                {
                    infoList.Add(info);
                    Console.Write("获取Excle文件:" + info.FullName+ "\n");
                }
            }

            return infoList;
        }

        /// <summary>
        /// 格式化Json字符串
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private string ConvertJsonString(string str)
        {
            //格式化json字符串
            JsonSerializer serializer = new JsonSerializer();
            TextReader tr = new StringReader(str);
            JsonTextReader jtr = new JsonTextReader(tr);
            object obj = serializer.Deserialize(jtr);
            if (obj != null)
            {
                StringWriter textWriter = new StringWriter();
                JsonTextWriter jsonWriter = new JsonTextWriter(textWriter)
                {
                    Formatting = Formatting.Indented,
                    Indentation = 4,
                    IndentChar = ' '
                };
                serializer.Serialize(jsonWriter, obj);
                return textWriter.ToString();
            }
            else
            {
                return str;
            }
        }
    }
}
*/
