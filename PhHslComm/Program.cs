// See https://aka.ms/new-console-template for more information


using HslCommunication;
using HslCommunication.LogNet;
using HslCommunication.Profinet.Omron;
using System.Net.Sockets;
using System.Reflection;
using System.Threading;
using System.Timers;
using Timer = System.Timers.Timer;
using System.Net;
using Newtonsoft.Json.Linq;
using HslCommunication.MQTT;
using HslCommunication.Profinet.Siemens;
using HslCommunication.Profinet.OpenProtocol;
using NPOI.XSSF.UserModel;
using static PhHslComm.UserStruct;
using Arp.Plc.Gds.Services.Grpc;
using Opc.Ua;
using Microsoft.Extensions.Logging;
using Google.Protobuf.WellKnownTypes;
using NPOI.POIFS.Crypt.Dsig;
using Grpc.Core;
using static Arp.Plc.Gds.Services.Grpc.IDataAccessService;
using Grpc.Net.Client;
using static PhHslComm.GrpcTool;
using static PhHslComm.OmronComm;
using System.Text;
using HslCommunication.CNC.Fanuc;
using System.Security.Claims;
using NPOI.Util;
using System.Text.RegularExpressions;
using SixLabors.ImageSharp;
using NPOI.SS.Formula.Functions;
using HslCommunication.Instrument.DLT;
using SixLabors.ImageSharp.Processing;
using System.Drawing;
using System.Net.NetworkInformation;


namespace PhHslComm
{
    class Program
    {
        /// <summary>
        /// app初始化
        /// </summary>

        // 创建日志
        const string logsFile = ("/opt/plcnext/apps/FengZhuangAppLogs.txt");
        //const string logsFile = "D:\\2024\\Work\\12-冠宇数采项目\\ReadFromStructArray\\GuanYu";
        public static ILogNet logNet = new LogNetSingle(logsFile);

        //创建Grpc实例
        public static GrpcTool grpcToolInstance = new GrpcTool();

        //设置grpc通讯参数
        public static CallOptions options1 = new CallOptions(
                new Metadata {
                        new Metadata.Entry("host","SladeHost")
                },
                DateTime.MaxValue,
                new CancellationTokenSource().Token);
        public static IDataAccessServiceClient grpcDataAccessServiceClient = null;
       
        //创建ASCII 转换API实例
        public static ToolAPI tool = new ToolAPI();

        //CIP Client实例化 
        public static OmronComm omronClients = new OmronComm();
        static int clientNum = 7; //CIP的Client最大32，开启7个CIP Client    
        public static OmronConnectedCipNet[] _cip = new OmronConnectedCipNet[clientNum];
        public static OperateResult ret;

        //创建三个线程            
        static int thrNum = 3;  //开启三个线程
        static Thread[] thr = new Thread[thrNum];

        //创建nodeID字典 (读取XML用）
        public static Dictionary<string, string> nodeidDictionary;

        //读取Excel用
        static ReadExcel readExcel = new ReadExcel();

        #region 从Excel解析来的数据实例化
        //设备信息里的离散数组数据
        static DeviceInfoConSturct1_CIP[] Auto_Process;
        static DeviceInfoConSturct1_CIP[] Clear_Manual;

        //设备信息里的离散结构体数据和数组数据
        static DeviceInfoConSturct1_CIP[] Battery_Memory;
        static DeviceInfoConSturct1_CIP[] BarCode;
        static DeviceInfoConSturct1_CIP[] EarCode;

        //六大工位数据
        static StationInfoStruct_CIP[] chongmo;
        static StationInfoStruct_CIP[] reya;
        static StationInfoStruct_CIP[] dingfeng;
        static StationInfoStruct_CIP[] zuojiaofeng;
        static StationInfoStruct_CIP[] youjiaofeng;
        static StationInfoStruct_CIP[] cefeng;

        //非报警信号（1000ms）
        static OneSecInfoStruct_CIP[] Sys_Manual;
        static OneSecInfoStruct_CIP[] Production_statistics;
        static OneSecInfoStruct_CIP[] Cutterused_statistics;
        static OneSecInfoStruct_CIP[] Y6;
        static OneSecInfoStruct_CIP[] Manual_Andon;

        //报警信号 （1000ms）
        static OneSecInfoStruct_CIP[] Vacuum_Alarm;
        static OneSecInfoStruct_CIP[] Senor_Alarm;
        static OneSecInfoStruct_CIP[] Motor_POTLimit_Err;
        static OneSecInfoStruct_CIP[] Motor_NOTLimit_Err;
        static OneSecInfoStruct_CIP[] Motor_Prevent_Promt;
        static OneSecInfoStruct_CIP[] Exception_information;
        static OneSecInfoStruct_CIP[] Temperature_Alarm;
        static OneSecInfoStruct_CIP[] Cylinder_Reset_Promt;
        static OneSecInfoStruct_CIP[] Motor_Reset_Promt;
        static OneSecInfoStruct_CIP[] Stopstate_Error;
        static OneSecInfoStruct_CIP[] Grating_Error;
        static OneSecInfoStruct_CIP[] Scapegoat_Error;
        static OneSecInfoStruct_CIP[] Door_Error;
        static OneSecInfoStruct_CIP[] out_power;

        //点位名
        static OneSecPointNameStruct_IEC oneSecPointNameStruct_IEC = new OneSecPointNameStruct_IEC();

        // 设备总览
        static DeviceInfoStruct_IEC[] deviceInfoStruct_IEC;

        #endregion

        // CIP连接状态
        static OperateResult retConnect;


        // 时间变量
        public static DateTime nowDisplay = DateTime.Now;
        static void Main(string[] args)
        {

            int stepNumber = 10;

            while (true)
            {
                switch (stepNumber)
                {

                    case 10:
                        {
                            /// <summary>
                            /// 执行初始化
                            /// </summary>


                            #region 读取Excel 

                            logNet.WriteError(nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff") + "App Start");

                            //string excelFilePath = Directory.GetCurrentDirectory() + "\\HGFZData.xlsx";     //PC端测试路径
                            string excelFilePath = "/opt/plcnext/apps/HGFZData.xlsx";                         //EPC存放路径

                            XSSFWorkbook excelWorkbook = readExcel.connectExcel(excelFilePath);

                            //Console.WriteLine("ExcelWorkbook read {0}", excelWorkbook != null ? "success" : "fail");
                            logNet.WriteError(nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff") + "  :ExcelWorkbook read ", excelWorkbook != null ? "success" : "fail");

                            //if (excelWorkbook != null)
                            //{
                            //    //Console.WriteLine("excelWorkbook reasd success");
                            //    logNet.WriteInfo("excelWorkbook reasd success");
                            //}
                            //else
                            //{
                            //    //Console.WriteLine("excelWorkbook reasd fail");
                            //    logNet.WriteError("excelWorkbook reasd fail");
                            //}

                            #endregion


                            #region 从xml获取nodeid，Grpc发送到对应变量时使用，注意xml中的别名要和对应类的属性名一致 
                            try
                            {
                                const string filePath = "/opt/plcnext/apps/GrpcSubscribeNodes.xml";             //EPC中存放的路径  
                                //const string filePath = "D:\\2024\\Work\\12-冠宇数采项目\\ReadFromStructArray\\FengZhuang_EIP\\PhHslComm\\GrpcSubscribeNodes\\GrpcSubscribeNodes.xml";  //PC中存放的路径 

                                nodeidDictionary = grpcToolInstance.getNodeIdDictionary(filePath);  //将xml中的值写入字典中
                                logNet.WriteError(nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff") + "  :NodeID read successfully");

                            }
                            catch (Exception e)
                            {
                                logNet.WriteError("Error:" + e);
                                logNet.WriteError(nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff") + "  :NodeID read failed");
                            }

                            #endregion


                            #region 将readExcel变量中的值，写入对应的实例化结构体中

                            // 六大工位（100ms)
                            chongmo = readExcel.ReadStationInfo_Excel(excelWorkbook, "加工工位（冲膜）");
                            reya = readExcel.ReadStationInfo_Excel(excelWorkbook, "加工工位（热压）");
                            dingfeng = readExcel.ReadStationInfo_Excel(excelWorkbook, "加工工位（顶封）");
                            zuojiaofeng = readExcel.ReadStationInfo_Excel(excelWorkbook, "加工工位（左角封）");
                            youjiaofeng = readExcel.ReadStationInfo_Excel(excelWorkbook, "加工工位（右角封）");
                            cefeng = readExcel.ReadStationInfo_Excel(excelWorkbook, "加工工位（侧封）");

                            // 非报警信号（1000ms）
                            Sys_Manual = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "功能开关");
                            Production_statistics = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "生产统计");
                            Cutterused_statistics = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "寿命管理");
                            Y6 = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "OEE");
                            Manual_Andon = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "OEE(2)");

                            // 报警信号（1000ms)
                            Vacuum_Alarm = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "Vacuum_Alarm");
                            Senor_Alarm = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "Senor_Alarm");
                            Motor_POTLimit_Err = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "Motor_POTLimit_Err");
                            Motor_NOTLimit_Err = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "Motor_NOTLimit_Err");
                            Motor_Prevent_Promt = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "Motor_Prevent_Promt");
                            Exception_information = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "Exception_information");
                            Temperature_Alarm = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "Temperature_Alarm");
                            Cylinder_Reset_Promt = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "Cylinder_Reset_Promt");
                            Motor_Reset_Promt = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "Motor_Reset_Promt");
                            Stopstate_Error = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "Stopstate_Error");
                            Grating_Error = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "Grating_Error");
                            Scapegoat_Error = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "Scapegoat_Error");
                            Door_Error = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "Door_Error");
                            out_power = readExcel.ReadOneSecInfo_Excel(excelWorkbook, "out_power");

                            // 设备信息（100ms）
                            Auto_Process = readExcel.ReadOneDeviceInfoConSturct1Info_Excel(excelWorkbook, "设备信息", 3);
                            Clear_Manual = readExcel.ReadOneDeviceInfoConSturct1Info_Excel(excelWorkbook, "设备信息", 5);
                            Battery_Memory = readExcel.ReadOneDeviceInfoConSturct1Info_Excel(excelWorkbook, "设备信息", 4);
                            BarCode = readExcel.ReadOneDeviceInfoConSturct2Info_Excel(excelWorkbook, "设备信息", 6);
                            EarCode = readExcel.ReadOneDeviceInfoConSturct2Info_Excel(excelWorkbook, "设备信息", 7);


                            #endregion


                            #region CIP连接   

                            //for (int i = 0; i < clientNum; i++)
                            //{
                            //    if (_cip[i] == null)
                            //    {
                            //        _cip[i] = new OmronConnectedCipNet("192.168.1.31");  //  填写欧姆龙PLC的IP地址
                            //        retIn = _cip[i].ConnectServer();
                            //        //logNet.WriteInfo("num " + i.ToString() + retIn.IsSuccess? "success" : "fail");
                            //        Console.WriteLine("num {0} connect: {1})!", i, retIn.IsSuccess? "success" : "fail");                                   
                            //    }
                            //    else
                            //    {
                            //        if(retIn.ErrorCode == 0)
                            //        {
                            //            Console.WriteLine("Connect open");
                            //        }
                            //        else 
                            //        {
                            //            _cip[i].ConnectClose();
                            //            Console.WriteLine("Connect closed");
                            //            retIn = _cip[i].ConnectServer();
                            //            //logNet.WriteInfo("num " + i.ToString() + retIn.IsSuccess? "success" : "fail");
                            //            Console.WriteLine("num {0} connect: {1})!", i, retIn.IsSuccess ? "success" : "fail");
                            //        }

                            //    }

                            //}




                            #endregion

                            #region Grpc连接 （TO DO LIST 先检查 后创建）

                            var udsEndPoint = new UnixDomainSocketEndPoint("/run/plcnext/grpc.sock");
                            var connectionFactory = new UnixDomainSocketConnectionFactory(udsEndPoint);

                            //grpcDataAccessServiceClient
                            var socketsHttpHandler = new SocketsHttpHandler
                            {
                                ConnectCallback = connectionFactory.ConnectAsync
                            };
                            try
                            {
                                GrpcChannel channel = GrpcChannel.ForAddress("http://localhost", new GrpcChannelOptions  // Create a gRPC channel to the PLCnext unix socket
                                {
                                    HttpHandler = socketsHttpHandler
                                });
                                grpcDataAccessServiceClient = new IDataAccessService.IDataAccessServiceClient(channel);// Create a gRPC client for the Data Access Service on that channel
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("ERRO: {0}", e);
                                //logNet.WriteError("Grpc connect failed!");
                            }
                            #endregion


                            // 单独设备总览表 + 发送
                            deviceInfoStruct_IEC = readExcel.ReadDeviceInfo_Excel(excelWorkbook, "封装设备总览");


                            var listWriteItem = new List<WriteItem>();
                            WriteItem[] writeItems = new WriteItem[] { };
                            try
                            {
                                listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["DeviceInfo"], Arp.Type.Grpc.CoreType.CtStruct, deviceInfoStruct_IEC[0]));
                                var writeItemsArray = listWriteItem.ToArray();
                                var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                                bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                            }
                            catch (Exception e)
                            {
                                //Console.WriteLine("ERRO: {0}，{1}", e, nodeidDictionary.GetValueOrDefault("DeviceInfo"));
                                logNet.WriteError(nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff") + "  DeviceInfo ERRO", e.ToString());
                            }
                            listWriteItem.Clear();


                            #region 1000ms数据的点位名

                            // 功能安全、生产统计和寿命管理 的点位名发送
                            omronClients.ReadandSendPointName(Sys_Manual, oneSecPointNameStruct_IEC, 200, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient,options1);   //功能开关 
                            omronClients.ReadandSendPointName(Production_statistics, oneSecPointNameStruct_IEC, 20, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);  //生产统计
                            omronClients.ReadandSendPointName(Cutterused_statistics, oneSecPointNameStruct_IEC, 36, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);  //寿命管理

                            //将Y6 和Manual_Andon的点位名拼成一个 string[]数组后，再发送 对应OEE表格
                            var stringnumber = Y6.Length + Manual_Andon.Length;
                            var OEEPointName = new string[stringnumber];
                            for (int i = 0; i < Y6.Length; i++)
                            {
                                OEEPointName[i] = Y6[i].varAnnotation;
                            }
                            for (int i = 0; i < Manual_Andon.Length; i++)
                            {
                                OEEPointName[i + Y6.Length] = Manual_Andon[i].varAnnotation;
                            }
                            omronClients.ReadandSendPointName(OEEPointName, oneSecPointNameStruct_IEC, 20, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);  //OEE


                            //报警信号的点位名拼成一个 string[]数组，再发送
                            stringnumber = Vacuum_Alarm.Length + Senor_Alarm.Length + Motor_POTLimit_Err.Length + Motor_NOTLimit_Err.Length
                                            + Motor_Prevent_Promt.Length + Exception_information.Length + Temperature_Alarm.Length
                                            + Cylinder_Reset_Promt.Length + Motor_Reset_Promt.Length + Stopstate_Error.Length + Grating_Error.Length
                                            + Scapegoat_Error.Length + Door_Error.Length + out_power.Length;
                            var AlarmPointName = new string[stringnumber];

                            List<OneSecInfoStruct_CIP[]> alarmGroups = new List<OneSecInfoStruct_CIP[]> {
                                Vacuum_Alarm, Senor_Alarm, Motor_POTLimit_Err, Motor_NOTLimit_Err, Motor_Prevent_Promt,
                                Exception_information,  Temperature_Alarm, Cylinder_Reset_Promt,
                                Motor_Reset_Promt, Stopstate_Error, Grating_Error, Scapegoat_Error, Door_Error, out_power
                            };
                            int index = 0;
                            foreach (var alarmGroup in alarmGroups)
                            {
                                foreach (var alarm in alarmGroup)
                                {
                                    AlarmPointName[index++] = alarm.varAnnotation;
                                }
                            }

                            omronClients.ReadandSendPointName(AlarmPointName, oneSecPointNameStruct_IEC, 2000, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);  //报警信号

                            #endregion


                            #region 六大工位的点位名
                            omronClients.ReadandSendPointName(chongmo, oneSecPointNameStruct_IEC, 8, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);  //冲膜
                            omronClients.ReadandSendPointName(dingfeng, oneSecPointNameStruct_IEC, 8, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);  //顶封
                            omronClients.ReadandSendPointName(reya, oneSecPointNameStruct_IEC, 29, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);  //热压
                            omronClients.ReadandSendPointName(zuojiaofeng, oneSecPointNameStruct_IEC, 8, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);  //左角封
                            omronClients.ReadandSendPointName(youjiaofeng, oneSecPointNameStruct_IEC, 8, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);  //右角封
                            omronClients.ReadandSendPointName(cefeng, oneSecPointNameStruct_IEC, 8, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);  //侧封

                            #endregion
                            stepNumber = 20;
                        }

                        break;

                    case 20:
                        {
                            #region CIP连接   

                            for (int i = 0; i < clientNum; i++)
                            {
                                if (_cip[i] == null)
                                {
                                    _cip[i] = new OmronConnectedCipNet("192.168.1.31");  //  填写欧姆龙PLC的IP地址
                                    var retConnect = _cip[i].ConnectServer();
                                    //logNet.WriteInfo("num " + i.ToString() + retIn.IsSuccess? "success" : "fail");
                                    Console.WriteLine("num {0} connect: {1})!", i, retConnect.IsSuccess ? "success" : "fail");
                                }
                                else
                                {
                                    if (retConnect.ErrorCode == 0)
                                    {
                                        Console.WriteLine("Connect open");
                                    }
                                    else
                                    {
                                        _cip[i].ConnectClose();
                                        Console.WriteLine("Connect closed");
                                        retConnect = _cip[i].ConnectServer();
                                        //logNet.WriteInfo("num " + i.ToString() + retIn.IsSuccess? "success" : "fail");
                                        Console.WriteLine("num {0} connect: {1})!", i, retConnect.IsSuccess ? "success" : "fail");
                                    }

                                }

                            }

                            #endregion
                            stepNumber = 90;
                        }
                        break;


                    case 90:
                        {
                            #region  线程初始化
                            //读1000ms数据
                            thr[0] = new Thread(() =>
                            {
                                var cip = _cip[0];
                                var listWriteItem = new List<WriteItem>();
                                WriteItem[] writeItems = new WriteItem[] { };
                                int lastIndex = out_power[out_power.Length - 1].varIndex;
                                bool[] Alarm_Data = new bool[lastIndex + 1];

                                while (true)
                                {
                                    TimeSpan start = new TimeSpan(DateTime.Now.Ticks);

                                    //读取报警信号总数组
                                    OperateResult<bool> ret = cip.ReadBool("Y6[5]");
                                    if (ret.IsSuccess)
                                    {
                                        if (ret.Content == true)
                                        {
                                            //读取Upload_Err总数组

                                            OperateResult<bool[]> retbool = cip.ReadBool(Vacuum_Alarm[0].varName, (ushort)(lastIndex + 1));
                                            if (ret.IsSuccess)
                                            {
                                                Alarm_Data = retbool.Content;
                                            }
                                            else
                                            {
                                                Console.WriteLine(Vacuum_Alarm[0].varName + "Array Read Failed");
                                            }

                                            //发送给IEC的
                                            var IECAlarmNumber = 2000;
                                            var AlarmValue = new bool[IECAlarmNumber];  //与IEC对应 2000 每次都清空
                                            var StartIndex = 0;
                                            var ArrayIndex = 0;

                                            //List里面元素的顺序不可以改变，与点位名的List要对应
                                            List<OneSecInfoStruct_CIP[]> alarmGroups = new List<OneSecInfoStruct_CIP[]> {
                                                    Vacuum_Alarm, Senor_Alarm, Motor_POTLimit_Err, Motor_NOTLimit_Err, Motor_Prevent_Promt,
                                                    Exception_information, Temperature_Alarm, Cylinder_Reset_Promt,
                                                    Motor_Reset_Promt, Stopstate_Error, Grating_Error, Scapegoat_Error, Door_Error, out_power
                                                };

                                            foreach (var alarmGroup in alarmGroups)
                                            {
                                                StartIndex = alarmGroup[0].varIndex;
                                                Array.Copy(Alarm_Data, StartIndex, AlarmValue, ArrayIndex, alarmGroup.Length);
                                                ArrayIndex += alarmGroup.Length;

                                            }

                                            #region Grpc发送数据
                                            try
                                            {
                                                listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["Alarm"], Arp.Type.Grpc.CoreType.CtArray, AlarmValue));
                                                var writeItemsArray = listWriteItem.ToArray();
                                                var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                                                bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                                            }
                                            catch (Exception e)
                                            {

                                                Console.WriteLine("ERRO: {0}，{1}", e, nodeidDictionary.GetValueOrDefault("Alarm"));
                                            }
                                            listWriteItem.Clear();
                                            #endregion 

                                        }
                                        else
                                        {
                                            //logNet.WriteInfo("No Warning");
                                            Console.WriteLine("No Warning");
                                        }

                                    }
                                    else
                                    {
                                        //logNet.WriteError("Y6[5] read failed");
                                        Console.WriteLine("Y6[5] read failed");
                                    }


                                    //读取非报警信号 Sys_Manual、Production_statistics、Cutterused_statistics的顺序不可以改变，与点位名发送一致
                                    omronClients.ReadandSendOneSecData(Sys_Manual, cip, 200, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);
                                    omronClients.ReadandSendOneSecData(Production_statistics, cip, 20, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);
                                    omronClients.ReadandSendOneSecData(Cutterused_statistics, cip, 36, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);

                                    bool[] Y6_temp = omronClients.ReadOneSecData(Y6, cip);
                                    bool[] Manual_Andon_temp = omronClients.ReadOneSecData(Manual_Andon, cip);

                                    var IECOEENumber = 20;
                                    var OEEValue = new bool[IECOEENumber];  //与IEC对应 20
                                    var OEEIndex = 0;

                                    //将Y6_temp和Manual_Andon 拼成一个OEE数组
                                    Array.Copy(Y6_temp, 0, OEEValue, OEEIndex, Y6_temp.Length);
                                    OEEIndex += Y6_temp.Length;
                                    Array.Copy(Manual_Andon_temp, 0, OEEValue, OEEIndex, Manual_Andon_temp.Length);

                                    #region Grpc发送数据给IEC
                                    try
                                    {
                                        listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["OEE"], Arp.Type.Grpc.CoreType.CtArray, OEEValue));
                                        var writeItemsArray = listWriteItem.ToArray();
                                        var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                                        bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                                    }
                                    catch (Exception e)
                                    {

                                        Console.WriteLine("ERRO: {0}", e, nodeidDictionary.GetValueOrDefault("OEE"));
                                    }
                                    listWriteItem.Clear();
                                    #endregion



                                    //计算从开始读到读完的时间
                                    TimeSpan end = new TimeSpan(DateTime.Now.Ticks);
                                    DateTime nowDisplay = DateTime.Now;
                                    TimeSpan dur = (end - start).Duration();

                                    //logNet.WriteInfo("Thread ReadOnceSecInfo read time : " + (dur.TotalMilliseconds).ToString());
                                    Console.WriteLine("Thread ReadOnceSecInfo read time:{0} read Duration:{1}", nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff"), dur.TotalMilliseconds);

                                    if (dur.TotalMilliseconds < 100)
                                    {
                                        int sleepTime = 1000 - (int)dur.TotalMilliseconds;
                                        Thread.Sleep(sleepTime);
                                    }

                                }
                            });

                            //读六大工位的数据
                            thr[1] = new Thread(() =>
                            {
                                var cip = _cip[1];

                                while (true)
                                {
                                    TimeSpan start = new TimeSpan(DateTime.Now.Ticks);

                                    #region 先一起读上来一整个数组，触发信号来了再通过Grpc发送

                                    //读取报警数组的最后一个索引数，索引数+1 等于数组元素数量
                                    int lastIndex = cefeng[cefeng.Length - 1].varIndex;
                                    var Station_Data = new float[lastIndex + 1];

                                    OperateResult<float[]> ret = cip.ReadFloat(chongmo[0].varName, (ushort)(lastIndex + 1));  //读取整个加工工位数据
                                    if (ret.IsSuccess)
                                    {
                                        Station_Data = ret.Content;
                                    }
                                    else
                                    {
                                        Console.WriteLine(Vacuum_Alarm[0].varName + "Array Read Failed");
                                    }

                                    OperateResult<int> retIn = cip.ReadInt32("Auto_process[32]");
                                    if (retIn.IsSuccess)
                                    {
                                        if (retIn.Content < 50 && retIn.Content >= 20)
                                            omronClients.SendSubArray(reya, Station_Data, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);
                                    }
                                    else
                                    {
                                        logNet.WriteError("Auto_process[32] read failed");
                                        //Console.WriteLine("Auto_process[32] read failed");
                                    }

                                    retIn = cip.ReadInt32("Auto_process[18]");
                                    if (retIn.IsSuccess)
                                    {
                                        if (retIn.Content < 50 && retIn.Content >= 25)
                                            omronClients.SendSubArray(dingfeng, Station_Data, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);
                                    }
                                    else
                                    {
                                        logNet.WriteInfo("Auto_process[18] read failed ");
                                        //Console.WriteLine("Auto_process[18] read failed");
                                    }

                                    retIn = cip.ReadInt32("Auto_process[9]");
                                    if (retIn.IsSuccess)
                                    {
                                        if (retIn.Content >= 40 && retIn.Content < 70)
                                            omronClients.SendSubArray(chongmo, Station_Data, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);
                                    }
                                    else
                                    {
                                        logNet.WriteInfo("Auto_process[9] read failed ");
                                        //Console.WriteLine("Auto_process[9] read failed");
                                    }

                                    retIn = cip.ReadInt32("Auto_process[43]");
                                    if (retIn.IsSuccess)
                                    {
                                        if (retIn.Content < 45 && retIn.Content >= 30)
                                            omronClients.SendSubArray(zuojiaofeng, Station_Data, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);
                                    }
                                    else
                                    {
                                        logNet.WriteInfo("Auto_process[43] read failed ");
                                        //Console.WriteLine("Auto_process[43] read failed");
                                    }

                                    retIn = cip.ReadInt32("Auto_process[44]");
                                    if (retIn.IsSuccess)
                                    {
                                        if (retIn.Content < 45 && retIn.Content >= 30)
                                            omronClients.SendSubArray(youjiaofeng, Station_Data, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);
                                    }
                                    else
                                    {
                                        //logNet.WriteInfo("Auto_process[44] read failed ");
                                        Console.WriteLine("Auto_process[44] read failed");
                                    }


                                    retIn = cip.ReadInt32("Auto_process[47]");
                                    if (retIn.IsSuccess)
                                    {
                                        if (retIn.Content < 60 && retIn.Content >= 30)
                                            omronClients.SendSubArray(cefeng, Station_Data, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);
                                    }
                                    else
                                    {
                                        //logNet.WriteInfo("Auto_process[47] read failed ");
                                        Console.WriteLine("Auto_process[47] read failed");
                                    }
                                    #endregion

                                    TimeSpan end = new TimeSpan(DateTime.Now.Ticks);
                                    DateTime nowDisplay = DateTime.Now;
                                    TimeSpan dur = (start - end).Duration();

                                    //logNet.WriteInfo("Thread ReadStationInfo1 read time : " + (dur.TotalMilliseconds).ToString());
                                    Console.WriteLine("Thread ReadStationInfo1 read time:{0} read Duration:{1}", nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff"), dur.TotalMilliseconds);

                                    if (dur.TotalMilliseconds < 100)
                                    {
                                        int sleepTime = 100 - (int)dur.TotalMilliseconds;
                                        Thread.Sleep(sleepTime);
                                    }

                                }
                            });

                            //读74个工位的数据
                            thr[2] = new Thread(() =>
                            {
                                int NumberOfStation = 75;      //取1-74号工位

                                string tempstring;

                                StringBuilder[] sbAutoProcess = new StringBuilder[NumberOfStation];
                                StringBuilder[] sbClearManual = new StringBuilder[NumberOfStation];
                                StringBuilder[] sbBatteryMemory = new StringBuilder[NumberOfStation];
                                StringBuilder[] sbBarCode = new StringBuilder[NumberOfStation];
                                StringBuilder[] sbEarCode = new StringBuilder[NumberOfStation];

                                //StringBuilder数组初始化
                                for (int i = 0; i < NumberOfStation; i++)
                                {
                                    sbAutoProcess[i] = new StringBuilder();
                                    sbClearManual[i] = new StringBuilder();
                                    sbBatteryMemory[i] = new StringBuilder();
                                    sbBarCode[i] = new StringBuilder();
                                    sbEarCode[i] = new StringBuilder();
                                }

                                var listWriteItem = new List<WriteItem>();
                                WriteItem[] writeItems = new WriteItem[] { };

                                stringStruct[] sendStringtoIEC = new stringStruct[74];

                                while (true)
                                {
                                    TimeSpan start = new TimeSpan(DateTime.Now.Ticks);

                                    //清空数据缓存区
                                    for (int i = 0; i < NumberOfStation; i++)
                                    {
                                        sbAutoProcess[i].Clear();
                                        sbClearManual[i].Clear();
                                        sbBatteryMemory[i].Clear();
                                        sbBarCode[i].Clear();
                                        sbEarCode[i].Clear();
                                    }

                                    omronClients.ReadDeviceInfoConSturct1(Auto_Process, _cip[2], sbAutoProcess);
                                    omronClients.ReadDeviceInfoConSturct1(Clear_Manual, _cip[3], sbClearManual);
                                    omronClients.ReadDeviceInfoConSturct1(Battery_Memory, _cip[4], sbBatteryMemory);
                                    omronClients.ReadDeviceInfoConSturct1(BarCode, _cip[5], sbBarCode);
                                    omronClients.ReadDeviceInfoConSturct1(EarCode, _cip[6], sbEarCode);




                                    //整合到74个string中 舍弃下标0 并行发送数据                         
                                    for (int i = 1; i < NumberOfStation; i++)
                                    {
                                        StringBuilder combinedString = new StringBuilder();   //每一行工位都是一个string

                                        combinedString.Append(sbAutoProcess[i]);
                                        if (sbBatteryMemory[i].Length == 0)
                                        {
                                            combinedString.Append(sbBatteryMemory[i] + " ,");

                                        }
                                        else
                                        {
                                            combinedString.Append(sbBatteryMemory[i]);
                                        }

                                        if (sbClearManual[i].Length == 0)
                                        {
                                            combinedString.Append(sbClearManual[i] + " ,");

                                        }
                                        else
                                        {
                                            combinedString.Append(sbClearManual[i]);
                                        }

                                        if (sbBarCode[i].Length == 0)
                                        {
                                            combinedString.Append(sbBarCode[i] + " ,");

                                        }
                                        else
                                        {
                                            combinedString.Append(sbBarCode[i]);
                                        }

                                        if (sbEarCode[i].Length == 0)
                                        {
                                            combinedString.Append(sbEarCode[i] + " ,");

                                        }
                                        else
                                        {
                                            combinedString.Append(sbEarCode[i]);
                                        }


                                        //combinedString.Append(sbBatteryMemory[i]);
                                        //combinedString.Append(sbClearManual[i]);
                                        //combinedString.Append(sbBarCode[i]);
                                        //combinedString.Append(sbEarCode[i]);
                                        //tempstring = combinedString.ToString().Replace(" ", "");
                                        tempstring = combinedString.ToString();
                                        sendStringtoIEC[i - 1].str = tempstring; //整合到结构体数组中 （一把子把74个工位的数据发送给IEC）

                                        #region Grpc发送数据给IEC

                                        if (i == 74)
                                        {
                                            try
                                            {
                                                listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["Station"], Arp.Type.Grpc.CoreType.CtArray, sendStringtoIEC));
                                                var writeItemsArray = listWriteItem.ToArray();
                                                var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                                                bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                                            }
                                            catch (Exception e)
                                            {
                                                Console.WriteLine("ERRO: {0}，{1}", e, nodeidDictionary.GetValueOrDefault(i.ToString()));
                                            }
                                            listWriteItem.Clear();
                                        }
                                        #endregion
                                    }


                                    TimeSpan end = new TimeSpan(DateTime.Now.Ticks);
                                    DateTime nowDisplay = DateTime.Now;
                                    TimeSpan dur = (start - end).Duration();

                                    //logNet.WriteInfo("Thread ReadDeviceInfo read time : " + (dur.TotalMilliseconds).ToString());
                                    Console.WriteLine("Thread ReadDeviceInfo read time:{0} read Duration:{1}", nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff"), dur.TotalMilliseconds);

                                    if (dur.TotalMilliseconds < 100)
                                    {
                                        int sleepTime = 100 - (int)dur.TotalMilliseconds;
                                        Thread.Sleep(sleepTime);
                                    }

                                }
                            });
                            #endregion

                            stepNumber = 100;

                        }
                        break;

                    

                    case 100:
                        {
                            if (thr[0].ThreadState == ThreadState.Unstarted && thr[1].ThreadState == ThreadState.Unstarted && thr[2].ThreadState == ThreadState.Unstarted)
                            {
                                try
                                {
                                    #region 开启三大数采线程

                                    thr[0].Start();  // 读1000ms数据
                                   
                                    thr[1].Start(); //读六大工位信息

                                    thr[2].Start();  //读设备信息                                       
                                    #endregion

                                }
                                catch
                                {
                                    Console.WriteLine("Thread quit");
                                    stepNumber = 1000;
                                    break;

                                }
                            }

                            IPStatus iPStatus;
                            iPStatus = _cip[0].IpAddressPing();  //判断与PLC的物理连接状态
                            
                            if (iPStatus != 0)
                            {
                                Console.WriteLine("Ping Omron PLC failed");

                            };

                            Thread.Sleep(100);

                        break;
                        }


                    case 1000:      //异常处理
                                    //信号复位
                                    //CIP连接断了


                        break;

                    case 10000:      //复位处理

                        break;


                }


            }

        }




    }
}