using System;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;

namespace excel2json
{
    /// <summary>
    /// 根据表头，生成C#类定义数据结构
    /// 表头使用三行定义：字段名称、字段类型、注释
    /// </summary>
    class CSTableGenerator
    {
        struct FieldDef
        {
            public string name;
            public string type;
            public string comment;
        }

        string mCode;

        public string code
        {
            get
            {
                return this.mCode;
            }
        }

        public CSTableGenerator(string excelName, ExcelLoader excel, string excludePrefix)
        {
            //-- 创建代码字符串
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("//");
            sb.AppendLine("// Auto Generated Code By excel2json");
            sb.AppendLine("// https://neil3d.gitee.io/coding/excel2json.html");
            sb.AppendLine("// 1. 每张表生成一个 Table");
            sb.AppendLine("// 2. 这个 class 是 partial 的,在 ConfigTableExtends 文件夹内扩展");
            sb.AppendLine("//");
            sb.AppendLine();
            sb.AppendFormat("// Generate From {0}.xlsx", excelName);
            sb.AppendLine();
            sb.AppendLine();

            if (excel.Sheets.Count > 0)
            {
                DataTable sheet = excel.Sheets[0];
                sb.Append(_generateTableCode(sheet, excelName, excludePrefix));
            }

            //for (int i = 0; i < excel.Sheets.Count; i++)
            //{
            //    DataTable sheet = excel.Sheets[i];
            //    sb.Append(_exportSheet(sheet, excludePrefix));
            //}

            sb.AppendLine();
            sb.AppendLine("// End of Auto Generated Code");

            mCode = sb.ToString();
        }

        private string _generateTableCode(DataTable sheet, string excelName, string excludePrefix)
        {
            //if (sheet.Columns.Count < 0 || sheet.Rows.Count < 2)
            //    return "";

            ////string sheetName = sheet.TableName;
            //if (excludePrefix.Length > 0 && excelName.StartsWith(excludePrefix))
            //    return "";

            //// get field list
            //List<FieldDef> fieldList = new List<FieldDef>();
            //DataRow typeRow = sheet.Rows[0];
            //DataRow commentRow = sheet.Rows[1];

            //foreach (DataColumn column in sheet.Columns)
            //{
            //    // 过滤掉包含指定前缀的列
            //    string columnName = column.ToString();
            //    if (excludePrefix.Length > 0 && columnName.StartsWith(excludePrefix))
            //        continue;

            //    FieldDef field;
            //    field.name = column.ToString();
            //    field.type = typeRow[column].ToString();
            //    field.comment = commentRow[column].ToString();

            //    fieldList.Add(field);
            //}

            // export as string
            StringBuilder sb = new StringBuilder();
            sb.Append("using System.Collections.Generic;\r\nusing WEngine.Runtime;\r\n");
            sb.AppendFormat("public partial class {0}Table\r\n{{", excelName);
            sb.AppendLine();

            sb.AppendFormat("\tprivate ResDictionary<int, {0}> m_config = null;\r\n", excelName);
            sb.AppendLine();
            sb.AppendFormat("\tpublic {0}Table()\r\n", excelName);
            sb.Append("\t{\r\n");
            sb.AppendFormat("\t\tm_config = new ResDictionary<int, {0}>();\r\n", excelName);
            sb.AppendFormat("\t\tm_config.Init(\"{0}\", ConvertID);\r\n", excelName);
            sb.Append("\t}\r\n");
            sb.AppendLine();

            sb.AppendFormat("\tprivate int ConvertID({0} config)\r\n", excelName);
            sb.Append("\t{\r\n");
            sb.Append("\t\treturn config.ID;\r\n");
            sb.Append("\t}\r\n");
            sb.AppendLine();

            sb.Append("\t// 根据ID获取配置\r\n");
            sb.AppendFormat("\tpublic {0} GetConfig(int id)\r\n", excelName);
            sb.Append("\t{\r\n");
            sb.AppendFormat("\t\tif (m_config.TryGetValue(id, out {0} config))\r\n", excelName);
            sb.Append("\t\t{\r\n");
            sb.Append("\t\t\treturn config;\r\n");
            sb.Append("\t\t}\r\n");
            sb.AppendLine();
            sb.Append("\t\treturn null;\r\n");
            sb.Append("\t}\r\n");
            sb.AppendLine();

            sb.Append("\t// 获取当前表所有数据\r\n");
            sb.AppendFormat("\tpublic Dictionary<int, {0}> GetAllConfigs()\r\n", excelName);
            sb.Append("\t{\r\n");
            sb.Append("\t\treturn m_config.Data;\r\n");
            sb.Append("\t}\r\n");

            sb.Append("}\r\n");
            return sb.ToString();
        }

        public void SaveToFile(string filePath, Encoding encoding)
        {
            //-- 保存文件
            using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                using (TextWriter writer = new StreamWriter(file, encoding))
                    writer.Write(mCode);
            }
        }
    }
}
