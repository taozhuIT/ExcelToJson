using System;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Security.AccessControl;

namespace ExcelToJson
{
    /// <summary>
    /// 工具类
    /// </summary>
    class Util
    {
        /// <summary>
        /// 判断字符串是否是数字
        /// </summary>
        public static bool IsNumber(string val_)
        {
            if (string.IsNullOrWhiteSpace(val_))
                return false;

            const string pattern = "^[0-9]*$";
            Regex rx = new Regex(pattern);
            return rx.IsMatch(val_);
        }

        /// <summary>
        /// 判断是否是浮点数
        /// </summary>
        /// <returns></returns>
        public static bool IsFloat(string val_)
        {
            float val;
            bool isOn = false;

            if (float.TryParse(val_, out val))
                isOn = true;

            return isOn;
        }

        /// <summary>
        /// 判断是否是布尔值
        /// </summary>
        /// <returns></returns>
        public static bool IsBool(string val_)
        {
            bool isOn = false;

            string valLower = val_.ToLower();
            // 这里加这么一句，因为在Excel里面的填写布尔值的单元格格式不是文本的话，获取到的excle的false或者true，会变成“真”或者“假”
            if (valLower.Equals("false") || valLower.Equals("true") || val_.Equals("真") || val_.Equals("假"))
                isOn = true;

            return isOn;
        }
    }
}
