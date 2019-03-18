using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Data;
using Newtonsoft.Json;
using System.Data.OleDb;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Text.RegularExpressions;

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
            Console.Write("请按任意键继续.......");
            Console.ReadLine();
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
                Dictionary<string, List<Dictionary<string, object>>> chartJsonStruct = new Dictionary<string, List<Dictionary<string, object>>>();
                
                for (int i = 0; i < tableList.Count; ++i)
                {
                    DataTable tableDt = tableList[i];
                    if (tableDt.Rows.Count < 2)
                        throw new Exception("表必须包含数据");

                    // 得到表中其中一个Table的json结构
                    List<Dictionary<string, object>> tableJsonStruct = null;
                    if (chartJsonStruct.ContainsKey(tableDt.TableName))
                    {
                        tableJsonStruct = chartJsonStruct[tableDt.TableName];
                    }
                    else
                    {
                        tableJsonStruct = new List<Dictionary<string, object>>();
                        chartJsonStruct.Add(tableDt.TableName, tableJsonStruct);
                    }

                    // table中的数据行
                    for (int n = 0; n < tableDt.Rows.Count; ++n)
                    {
                        // 0行是注释
                        if (n > 0)
                        {
                            DataRow headRow = tableDt.Rows[n];

                            // 第1行配置是json_key
                            if (n == 1)
                            {
                                keys.Clear();
                                for (int j = 0; j < headRow.ItemArray.Length; ++j)
                                {
                                    keys.Add(headRow.ItemArray[j].ToString());
                                }
                            }
                            else
                            {
                                // 得到表中其中一个Table中的一个数据类json的结构
                                Dictionary<string, object> itemStruct = new Dictionary<string, object>();
                                for (int x = 0; x < headRow.ItemArray.Length; ++x)
                                {
                                    string key = keys[x];

                                    object valObj = headRow.ItemArray[x];
                                    if (IsNumber(valObj.ToString()))
                                    {
                                        int val = int.Parse(headRow.ItemArray[x].ToString());
                                        itemStruct.Add(key, val);
                                    }
                                    else
                                    {
                                        string val = headRow.ItemArray[x].ToString();
                                        itemStruct.Add(key, val);
                                    }
                                }
                                tableJsonStruct.Add(itemStruct);
                            }
                        }
                    }
                    // ------------------------创建序列化Json的数据结构------------------------
                }

                Console.Write("创建  " + info.Name + "  Json数据结构完成\n");
                OnJsonSerializer(info, chartJsonStruct);
                Console.Write("--------转换:  " + info.Name + "  到Json文件完成--------\n\n");
            }
        }

        /// <summary>
        /// 序列化数据成json文件
        /// </summary>
        private void OnJsonSerializer(FileSystemInfo fileInfo_, Dictionary<string, List<Dictionary<string, object>>> chartJsonStruct_)
        {
            Console.Write("开始  "+ fileInfo_.Name+ "  文件数据json序列化\n");
            string creatPath = fileInfo_.FullName.Replace(@"\Conf", @"\Json").Replace(".xlsx", ".json");

            // 创建存放Json文件的文件夹
            int folderIndex = creatPath.LastIndexOf(@"\");
            string jsonFolderPath = creatPath.Substring(0, folderIndex);
            if (!Directory.Exists(jsonFolderPath))
                Directory.CreateDirectory(jsonFolderPath);

            // 序列化json数据
            string SerializeJson = JsonConvert.SerializeObject(chartJsonStruct_);
            
            // 创建json文件
            if (File.Exists(creatPath))
                File.Delete(creatPath);
            FileStream fs = File.Create(creatPath);

            // 写入数据
            // 格式json字符串（不格式话，显示就是一行显示所有数据）
            string convertJson = ConvertJsonString(SerializeJson);
            byte[] byteJson = Encoding.Default.GetBytes(convertJson);
            fs.Write(byteJson, 0, byteJson.Length);
            fs.Close();
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

        /// <summary>
        /// 判断字符串是否是数字
        /// </summary>
        public bool IsNumber(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            const string pattern = "^[0-9]*$";
            Regex rx = new Regex(pattern);
            return rx.IsMatch(s);
        }
    }
}
