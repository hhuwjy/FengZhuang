using Microsoft.VisualBasic;
using Newtonsoft.Json;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using static Ph_CipComm_FengZhuang.UserStruct;

namespace Ph_CipComm_FengZhuang
{

    public class ReadExcel
    {
       
        public XSSFWorkbook connectExcel(string excelFilePath)
        {
            XSSFWorkbook xssWorkbook = null;

            if (!File.Exists(excelFilePath))
            {
                Console.WriteLine(excelFilePath + ": 读取的文件不存在");
                return xssWorkbook;
            }

           

            try {
                using (FileStream stream = new FileStream(excelFilePath, FileMode.Open))
                {
                    stream.Position = 0;
                    xssWorkbook = new XSSFWorkbook(stream);
                    stream.Close();
                }
            }
            catch (Exception )
            {
                return xssWorkbook;
                throw;

            }


            return xssWorkbook;

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="excelFilePath"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public StationInfoStruct_CIP[] ReadStationInfo_Excel(XSSFWorkbook xssWorkbook, string sheetName)
        {
            DataTable dtTable = new DataTable();
            List<string> rowList = new List<string>();
            
           
            //sheet = xssWorkbook.GetSheetAt(0);
            ISheet sheet = xssWorkbook.GetSheet(sheetName);
            if (sheet == null)
            {
                Console.WriteLine(sheetName+ "页不存在");
                return null;

            }


            IRow headerRow = sheet.GetRow(0);
            int cellCount = headerRow.LastCellNum;

     

            List<StationInfoStruct_CIP> retList = new List<StationInfoStruct_CIP>();


            for (int j = 0; j < cellCount; j++)
            {
                ICell cell = headerRow.GetCell(j);
                if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
                {
                    dtTable.Columns.Add(cell.ToString());
                }
            }
            for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null) continue;
                if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;

                string str = Convert.ToString(row.GetCell(1));
                if (string.IsNullOrEmpty(str) || string.IsNullOrWhiteSpace(str)) continue;


                var v = new StationInfoStruct_CIP();
                for (int j = row.FirstCellNum; j < cellCount; j++)
                {
                    if (row.GetCell(j) != null)
                    {
                        if (!string.IsNullOrEmpty(row.GetCell(j).ToString()) && !string.IsNullOrWhiteSpace(row.GetCell(j).ToString()))
                        {
                            v.stationName = Convert.ToString(sheetName);
                            if (j == 1)
                            {
                                v.varName = Convert.ToString(row.GetCell(j));
                                if (!(string.IsNullOrEmpty(v.varName) || string.IsNullOrWhiteSpace(v.varName)))
                                {
                                    Regex r = new Regex(@"(?i)(?<=\[)(.*)(?=\])");//中括号[]
                                    var ms = r.Matches(v.varName);
                                    if (ms.Count > 0)
                                        v.varIndex = Convert.ToInt32(ms.ToArray()[0].Value);
                                }

                            }
                            else if (j == 2)
                            {
                                v.varAnnotation = Convert.ToString(row.GetCell(j));

                            }
                            else if (j == 3)
                            {
                                v.varType = Convert.ToString(row.GetCell(j));

                            }
                        }
                    }
                }

                retList.Add(v);
            }

            return retList.ToArray(); ;
        }


        //从Excel中读取1秒的数据信息
        public OneSecInfoStruct_CIP[] ReadOneSecInfo_Excel(XSSFWorkbook xssWorkbook, string sheetName)
        {
            

            DataTable dtTable = new DataTable();
            List<string> rowList = new List<string>();

             
           
            //sheet = xssWorkbook.GetSheetAt(0);
            ISheet sheet = xssWorkbook.GetSheet(sheetName);
            if (sheet == null)
            {
                Console.WriteLine(sheetName + "页不存在");
                return null;

            }


            IRow headerRow = sheet.GetRow(0);
            int cellCount = headerRow.LastCellNum;

            List<OneSecInfoStruct_CIP> retList = new List<OneSecInfoStruct_CIP>();


            for (int j = 0; j < cellCount; j++)
            {
                ICell cell = headerRow.GetCell(j);
                if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
                {
                    dtTable.Columns.Add(cell.ToString());
                }
            }
            for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null) continue;
                if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;

                string str = Convert.ToString(row.GetCell(1));
                if (string.IsNullOrEmpty(str) || string.IsNullOrWhiteSpace(str)) continue;


                var v = new OneSecInfoStruct_CIP();

                for (int j = row.FirstCellNum; j < cellCount; j++)
                {
                    if (j == 1)
                    {
                        v.varName = Convert.ToString(row.GetCell(j).StringCellValue).Trim();
                        if (!(string.IsNullOrEmpty(v.varName) || string.IsNullOrWhiteSpace(v.varName)))
                        {
                            Regex r = new Regex(@"(?i)(?<=\[)(.*)(?=\])");//中括号[]
                            var ms = r.Matches(v.varName);
                            if (ms.Count > 0)
                                v.varIndex = Convert.ToInt16(ms.ToArray()[0].Value);
                        }
                    }
                    else if (j == 2)
                    {
                        v.varAnnotation = Convert.ToString(row.GetCell(j));

                        //if (string.IsNullOrWhiteSpace(v.varAnnotation) || string.IsNullOrEmpty(v.varAnnotation))
                        //{
                        //    v.varIndex = -1;
                        //}
                        //else
                        //{
                        //    v.varIndex = i - 1;
                        //}

                        //varIndex
                       
                       

                    }
                    else if (j == 3)
                    {
                        v.varType = Convert.ToString(row.GetCell(j));
                    }
                }
                retList.Add(v);
            }

            return retList.ToArray();
        }

        //从Excel中读取DeviceInfo的数据信息1
        public DeviceInfoConSturct1_CIP[] ReadOneDeviceInfoConSturct1Info_Excel(XSSFWorkbook xssWorkbook, string sheetName,int columnNumber)
        {
            

            DataTable dtTable = new DataTable();
            List<string> rowList = new List<string>();

             
            ISheet sheet = xssWorkbook.GetSheet(sheetName);
            if (sheet == null)
            {
                Console.WriteLine( sheetName + "页不存在");
                return null;

            }


            IRow headerRow = sheet.GetRow(0);
            int cellCount = headerRow.LastCellNum;

            List<DeviceInfoConSturct1_CIP> retList = new List<DeviceInfoConSturct1_CIP>();


            for (int j = 0; j < cellCount; j++)
            {
                ICell cell = headerRow.GetCell(j);
                if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
                {
                    dtTable.Columns.Add(cell.ToString());
                }
            }
            for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null) continue;
                if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;

                string str = Convert.ToString(row.GetCell(columnNumber - 1));
                if (string.IsNullOrEmpty(str) || string.IsNullOrWhiteSpace(str)) continue;

                var v = new DeviceInfoConSturct1_CIP();

                for (int j = row.FirstCellNum; j < cellCount; j++)
                {
                    if (j == 0)
                    {
                        v.stationNumber = Convert.ToInt32(row.GetCell(j).NumericCellValue);
                    }
                    else if (j == 1)
                    {
                        v.stationName = Convert.ToString(row.GetCell(j));
                        

                    }
                    else if (j == (columnNumber-1))
                        {
                 
                        //varIndex
                        string temp = Convert.ToString(row.GetCell(j));
                        if (!(string.IsNullOrEmpty(temp) || string.IsNullOrWhiteSpace(temp)))
                        {
                            Regex r = new Regex(@"(?i)(?<=\[)(.*)(?=\])");//中括号[]
                            var ms = r.Matches(temp);
                            if (ms.Count > 0)                          
                            v.varIndex = Convert.ToInt16(ms.ToArray()[0].Value);
                        }

                        //varName
                        int index = temp.IndexOf('[');
                        if(index > -1)
                        v.varName = temp.Substring(0, index); 


                        //varType
                        temp = Convert.ToString(headerRow.GetCell(j));
                        if (!(string.IsNullOrEmpty(temp) || string.IsNullOrWhiteSpace(temp)))
                        {
                            Regex r = new Regex(@"\((\w+)\)");
                            var ms = r.Matches(getNewString(temp));
                            if( ms.Count>0)
                            v.varType = ms.ToArray()[0].Groups[1].Value;

                        }

                            

                        

                    }
                   
                    
                    

                }
                retList.Add(v);
            }

            return retList.ToArray();
        }

        public DeviceInfoConSturct1_CIP[] ReadOneDeviceInfoConSturct2Info_Excel(XSSFWorkbook xssWorkbook, string sheetName, int columnNumber)
        {


            DataTable dtTable = new DataTable();
            List<string> rowList = new List<string>();


            ISheet sheet = xssWorkbook.GetSheet(sheetName);
            if (sheet == null)
            {
                Console.WriteLine(sheetName + "页不存在");
                return null;

            }


            IRow headerRow = sheet.GetRow(0);
            int cellCount = headerRow.LastCellNum;

            List<DeviceInfoConSturct1_CIP> retList = new List<DeviceInfoConSturct1_CIP>();


            for (int j = 0; j < cellCount; j++)
            {
                ICell cell = headerRow.GetCell(j);
                if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
                {
                    dtTable.Columns.Add(cell.ToString());
                }
            }
            for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null) continue;
                if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;

                string str = Convert.ToString(row.GetCell(columnNumber - 1));
                if (string.IsNullOrEmpty(str) || string.IsNullOrWhiteSpace(str)) continue;

                var v = new DeviceInfoConSturct1_CIP();

                for (int j = row.FirstCellNum; j < cellCount; j++)
                {
                    if (j == 0)
                    {
                        v.stationNumber = Convert.ToInt32(row.GetCell(j).NumericCellValue);
                    }
                    else if (j == 1)
                    {
                        v.stationName = Convert.ToString(row.GetCell(j));


                    }
                    else if (j == (columnNumber - 1))
                    {

                        //varIndex
                        string temp = Convert.ToString(row.GetCell(j));
                        v.varIndex = -1;
                        //if (!(string.IsNullOrEmpty(temp) || string.IsNullOrWhiteSpace(temp)))
                        //{
                        //    Regex r = new Regex(@"(?i)(?<=\[)(.*)(?=\])");//中括号[]
                        //    var ms = r.Matches(temp);
                        //    if (ms.Count > 0)
                        //        v.varIndex = Convert.ToInt16(ms.ToArray()[0].Value);
                        //}

                        //varName
                        int index = temp.IndexOf('[');
                        if (index > -1)
                            v.varName = temp.Substring(0, index);


                        //varType
                        temp = Convert.ToString(headerRow.GetCell(j));
                        if (!(string.IsNullOrEmpty(temp) || string.IsNullOrWhiteSpace(temp)))
                        {
                            Regex r = new Regex(@"\((\w+)\)");
                            var ms = r.Matches(getNewString(temp));
                            if (ms.Count > 0)
                                v.varType = ms.ToArray()[0].Groups[1].Value;

                        }





                    }




                }
                retList.Add(v);
            }

            return retList.ToArray();
        }



        //从Excel中读取DeviceInfo的数据信息2
        // public DeviceInfoConSturct1_CIP[] ReadOneDeviceInfoDisSturct2Info_Excel(XSSFWorkbook xssWorkbook, string sheetName, int columnNumber)
        // {


        //    DataTable dtTable = new DataTable();
        //    List<string> rowList = new List<string>();


        //    ISheet sheet = xssWorkbook.GetSheet(sheetName);
        //    if (sheet == null)
        //    {
        //        Console.WriteLine(  sheetName + "页不存在");
        //        return null;

        //    }


        //    IRow headerRow = sheet.GetRow(0);
        //    int cellCount = headerRow.LastCellNum;

        //    //DeviceInfoDisStruct2_CIP[] ret = new DeviceInfoDisStruct2_CIP[sheet.LastRowNum];

        //    List< DeviceInfoDisStruct2_CIP > retList = new List<DeviceInfoDisStruct2_CIP>();

        //    for (int j = 0; j < cellCount; j++)
        //    {
        //        ICell cell = headerRow.GetCell(j);
        //        if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
        //        {
        //            dtTable.Columns.Add(cell.ToString());
        //        }
        //    }
        //    for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
        //    {
        //        IRow row = sheet.GetRow(i);
        //        if (row == null) continue;
        //        if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;

        //        var v = new DeviceInfoDisStruct2_CIP();

        //        string str = Convert.ToString(row.GetCell(columnNumber - 1));
        //        if (string.IsNullOrEmpty(str) || string.IsNullOrWhiteSpace(str)) continue;

        //        for (int j = row.FirstCellNum; j < cellCount; j++)
        //        {

        //            if (j == (columnNumber - 1))
        //            {

        //                string temp = Convert.ToString(row.GetCell(j));                        

        //                //varName
        //                int indext = (temp).LastIndexOf('[') == (temp).IndexOf('[') ? -1 : (temp).LastIndexOf('[');

        //                if (indext > -1)
        //                {
        //                    v.varName = temp.Substring(0, indext);
        //                }
        //                else
        //                {
        //                    v.varName = temp;
        //                }

        //                //if (string.IsNullOrWhiteSpace(ret[i - 1].varAnnotation) || string.IsNullOrEmpty(ret[i - 1].varAnnotation))


        //                //varType
        //                temp = Convert.ToString(headerRow.GetCell(j));
        //                if (!(string.IsNullOrEmpty(temp) || string.IsNullOrWhiteSpace(temp)))
        //                {
        //                    Regex r = new Regex(@"\((\w+)\)");
        //                    var ms = r.Matches(getNewString(temp));
        //                    if (ms.Count > 0)
        //                        v.varType = ms.ToArray()[0].Groups[1].Value;

        //                }


        //            }
        //            else if (j == 0)
        //            {
        //                v.stationNumber = Convert.ToInt32(row.GetCell(j).NumericCellValue);
        //            }
        //            else if (j == 1)
        //            {
        //                v.stationName = Convert.ToString(row.GetCell(j));


        //            }

        //        }

        //        retList.Add(v);
        //    }

        //    return retList.ToArray();
        //}


        //[DllImport("kernel32.dll")]
        //public static extern IntPtr _lopen(string lpPathName, int iReadWrite);

        //[DllImport("kernel32.dll")]
        //public static extern bool CloseHandle(IntPtr hObject);
        //public const int OF_READWRITE = 2;
        //public const int OF_SHARE_DENY_NONE = 0x40;
        //public static readonly IntPtr HFILE_ERROR = new IntPtr(-1);



        /// <summary>
        /// 文件是否被打开
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        //public static bool IsFileOpen(string path)
        //{
        //    if (!File.Exists(path))
        //    {
        //        return false;
        //    }
        //    IntPtr vHandle = _lopen(path, OF_READWRITE | OF_SHARE_DENY_NONE);//windows Api上面有定义扩展方法
        //    if (vHandle == HFILE_ERROR)
        //    {
        //        return true;
        //    }
        //    CloseHandle(vHandle);
        //    return false;
        //}

        public static string getNewString(String Node)
        {
            String newNode = null;
            String allConvertNode = null;
            if (Node.Contains("（") && Node.Contains("）"))
            {
                newNode = Node.Replace("（", "(");
                allConvertNode = newNode.Replace("）", ")");
            }
            else if (!(Node.Contains("（")) && Node.Contains("）"))
            {
                allConvertNode = Node.Replace("）", ")");
            }
            else if (Node.Contains("（") && !(Node.Contains("）")))
            {
                newNode = Node.Replace("（", "(");
                allConvertNode = newNode;
            }
            else
            {
                allConvertNode = Node;
            }
            return allConvertNode;
        }




        //读取封装设备信息
        public DeviceInfoStruct_IEC[] ReadDeviceInfo_Excel(XSSFWorkbook xssWorkbook, string sheetName)
        {
            List<DeviceInfoStruct_IEC> deviceInfoStruct_IEC = new List<DeviceInfoStruct_IEC>();
            try
            {
                DataTable dtTable = new DataTable();
                List<string> rowList = new List<string>();
                ISheet sheet = xssWorkbook.GetSheet(sheetName.Trim());
                if (sheet == null)
                {
                    Console.WriteLine(sheetName + "页不存在");
                    return null;
                }

                IRow headerRow = sheet.GetRow(0);
                int cellCount = headerRow.LastCellNum;


                for (int j = 0; j < cellCount; j++)
                {
                    ICell cell = headerRow.GetCell(j);
                    if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
                    {
                        dtTable.Columns.Add(cell.ToString());
                    }
                }
                for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue;
                    if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;
                    DeviceInfoStruct_IEC v = new DeviceInfoStruct_IEC();

                    string str = Convert.ToString(row.GetCell(0));
                    if (string.IsNullOrEmpty(str) || string.IsNullOrWhiteSpace(str)) continue;

                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        if (row.GetCell(j) != null)
                        {
                            if (!string.IsNullOrEmpty(row.GetCell(j).ToString()) && !string.IsNullOrWhiteSpace(row.GetCell(j).ToString()))
                            {

                                if (j == 0)
                                {
                                    v.strDeviceName = string.IsNullOrEmpty(Convert.ToString(row.GetCell(j))) ? " " : Convert.ToString(row.GetCell(j));

                                }
                                else if (j == 1)
                                {
                                    v.strDeviceCode = string.IsNullOrEmpty(Convert.ToString(row.GetCell(j))) ? " " : Convert.ToString(row.GetCell(j));

                                }
                                else if (j == 2)
                                {
                                    v.strPLCType = string.IsNullOrEmpty(Convert.ToString(row.GetCell(j))) ? " " : Convert.ToString(row.GetCell(j));

                                }
                                else if (j == 3)
                                {
                                    v.strProtocol = string.IsNullOrEmpty(Convert.ToString(row.GetCell(j))) ? " " : Convert.ToString(row.GetCell(j));

                                }
                                else if (j == 4)
                                {
                                    v.strIPAddress = string.IsNullOrEmpty(Convert.ToString(row.GetCell(j))) ? " " : Convert.ToString(row.GetCell(j));

                                }
                                else if (j == 5)
                                {
                                    v.iPort = Convert.ToInt32 (row.GetCell(j).NumericCellValue);//这里超出int16的范围  

                                }
                                else if (j == 6)
                                {
                                    v.iStationCount = Convert.ToInt16(row.GetCell(j).NumericCellValue);

                                }
                            }
                        }
                    }
                    deviceInfoStruct_IEC.Add(v);
                }

            }
            catch (Exception)
            {
                throw;
            }
            return deviceInfoStruct_IEC.ToArray(); ;


        }



    }


    

}
