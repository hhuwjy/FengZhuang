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
using static Ph_CipComm_FengZhuang.UserStruct;
using Arp.Plc.Gds.Services.Grpc;
using Opc.Ua;
using Microsoft.Extensions.Logging;
using Google.Protobuf.WellKnownTypes;
using NPOI.POIFS.Crypt.Dsig;
using Grpc.Core;
using static Arp.Plc.Gds.Services.Grpc.IDataAccessService;
using Grpc.Net.Client;
using static Ph_CipComm_FengZhuang.GrpcTool;
using static Ph_CipComm_FengZhuang.OmronComm;
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


namespace Ph_CipComm_FengZhuang
{
    class Program
    {
        /// <summary>
        /// app初始化
        /// </summary>

        // 创建日志
        const string logsFile = ("/opt/plcnext/apps/FengZhuangAppLogs.txt");
        //const string logsFile = "D:\\2024\\Work\\12-冠宇数采项目\\ReadFromStructArray\\FengZhuang_EIP\\FengZhuangAppLogs.txt";
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
        //public static OperateResult ret;

        //创建三个线程            
        static int thrNum = 4;  //开启三个线程
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

        //（六大工位，生产统计，寿命管理 ，电芯码 、极耳码、OEE数据）数据暂存结构体
        static AllDataReadfromCIP allDataReadfromCIP = new AllDataReadfromCIP();

        // 设备总览
        static DeviceInfoStruct_IEC[] deviceInfoStruct_IEC;

        #endregion

        // CIP连接状态
        static OperateResult retConnect;


        // 时间变量
        public static DateTime nowDisplay = DateTime.Now;

        static void Main(string[] args)
        {

            int stepNumber = 5;

            List<WriteItem> listWriteItem = new List<WriteItem>();
            IDataAccessServiceReadSingleRequest dataAccessServiceReadSingleRequest = new IDataAccessServiceReadSingleRequest();

            bool isThreadOneRunning = false;
            bool isThreadTwoRunning = false;
            bool isThreadThreeRunning = false;

            #region 从xml获取nodeid，Grpc发送到对应变量时使用，注意xml中的别名要和对应类的属性名一致 
            try
            {
                const string filePath = "/opt/plcnext/apps/GrpcSubscribeNodes.xml";             //EPC中存放的路径  
                //const string filePath = "D:\\2024\\Work\\12-冠宇数采项目\\ReadFromStructArray\\FengZhuang_EIP\\Ph_CipComm_FengZhuang\\GrpcSubscribeNodes\\GrpcSubscribeNodes.xml";  //PC中存放的路径 

                nodeidDictionary = grpcToolInstance.getNodeIdDictionary(filePath);  //将xml中的值写入字典中
                logNet.WriteInfo(nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff") + "  :NodeID read successfully");

            }
            catch (Exception e)
            {
                logNet.WriteError("Error:" + e);
                logNet.WriteError(nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff") + "  :NodeID read failed");
            }
            #endregion


            while (true)
            {
                switch (stepNumber)
                {

                    case 5:
                        {
                            ///连接Grpc
                            #region Grpc连接 

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

                            stepNumber = 10;
                        }
                        break;



                    case 10:
                        {
                            /// <summary>
                            /// 执行初始化
                            /// </summary>

                         
                            #region 读取Excel 

                            logNet.WriteInfo(nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff") + " :App Start");

                            //string excelFilePath = Directory.GetCurrentDirectory() + "\\HGFZData.xlsx";     //PC端测试路径
                            string excelFilePath = "/opt/plcnext/apps/HGFZData.xlsx";                         //EPC存放路径

                            XSSFWorkbook excelWorkbook = readExcel.connectExcel(excelFilePath);

                            Console.WriteLine("ExcelWorkbook read {0}", excelWorkbook != null ? "success" : "fail");
                            if (excelWorkbook != null)
                            {
                                logNet.WriteInfo(nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff") + "  :ExcelWorkbook read success");
                            }
                            else
                            {
                                logNet.WriteError(nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff") + "  :ExcelWorkbook read fail");
                            }
                           


                            // 给IEC发送 Excel读取成功的信号
                            var tempFlag_finishReadExcelFile = true;
                            
                            listWriteItem.Clear();
                            listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["flag_finishReadExcelFile"], Arp.Type.Grpc.CoreType.CtBoolean, tempFlag_finishReadExcelFile));
                            if (grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, grpcToolInstance.ServiceWriteRequestAddDatas(listWriteItem.ToArray()), new IDataAccessServiceWriteResponse(), options1))
                            {
                                Console.WriteLine("{0}      flag_finishReadExcelFile写入IEC: success", DateTime.Now);
                            }
                            else
                            {
                                Console.WriteLine("{0}      flag_finishReadExcelFile写入IEC: fail", DateTime.Now);
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
                            Auto_Process = readExcel.ReadOneDeviceInfoConSturct1Info_Excel(excelWorkbook, "设备信息", "电芯记忆信号(DINT)");
                            Clear_Manual = readExcel.ReadOneDeviceInfoConSturct1Info_Excel(excelWorkbook, "设备信息", "电芯记忆清除按钮(BOOL)");
                            Battery_Memory = readExcel.ReadOneDeviceInfoConSturct1Info_Excel(excelWorkbook, "设备信息", "电芯记忆(BOOL)");
                            BarCode = readExcel.ReadOneDeviceInfoConSturct2Info_Excel(excelWorkbook, "设备信息", "电芯条码地址（REAL)");
                            EarCode = readExcel.ReadOneDeviceInfoConSturct2Info_Excel(excelWorkbook, "设备信息", "极耳码地址(REAL)");


                            #endregion


                            #region 单独设备总览表 + 发送
                            deviceInfoStruct_IEC = readExcel.ReadDeviceInfo_Excel(excelWorkbook, "封装设备总览");

                            //var listWriteItem = new List<WriteItem>();
                            //WriteItem[] writeItems = new WriteItem[] { };

                            listWriteItem.Clear();
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
                            //listWriteItem.Clear();

                            #endregion


                            #region 读取并发送1000ms数据的点位名

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


                            #region 读取并发送六大工位的点位名
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
                                _cip[i] = new OmronConnectedCipNet(deviceInfoStruct_IEC[0].strIPAddress);  //  填写欧姆龙PLC的IP地址
                                var retConnect = _cip[i].ConnectServer();
                                
                                Console.WriteLine("num {0} connect: {1})!", i, retConnect.IsSuccess ? "success" : "fail");


                                //if (_cip[i] == null)
                                //{
                                //    _cip[i] = new OmronConnectedCipNet(deviceInfoStruct_IEC[0].strIPAddress);  //  填写欧姆龙PLC的IP地址
                                //    var retConnect = _cip[i].ConnectServer();
                                //    logNet.WriteInfo("num " + i.ToString() + retIn.IsSuccess? "success" : "fail");
                                //    Console.WriteLine("num {0} connect: {1})!", i, retConnect.IsSuccess ? "success" : "fail");
                                //}
                                //else
                                //{
                                //    if (retConnect.ErrorCode == 0)
                                //    {
                                //        Console.WriteLine("Connect open");
                                //    }
                                //    else
                                //    {
                                //        _cip[i].ConnectClose();
                                //        Console.WriteLine("Connect closed");
                                //        retConnect = _cip[i].ConnectServer();
                                //        logNet.WriteInfo("num " + i.ToString() + retIn.IsSuccess? "success" : "fail");
                                //        Console.WriteLine("num {0} connect: {1})!", i, retConnect.IsSuccess ? "success" : "fail");
                                //    }

                                //}

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

                                while (isThreadOneRunning)
                                {
                                    TimeSpan start = new TimeSpan(DateTime.Now.Ticks);

                                    #region 读取并发送报警信号 （读总数组，进行拼接）
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
                                    #endregion


                                    #region 读取并发送非报警信号

                                    //读取非报警信号 Sys_Manual、Production_statistics、Cutterused_statistics的顺序不可以改变，与点位名发送一致
                                    omronClients.ReadandSendOneSecData(Sys_Manual, cip, ref  allDataReadfromCIP, 200, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);
                                    omronClients.ReadandSendOneSecData(Production_statistics, cip, ref allDataReadfromCIP, 20, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);
                                    omronClients.ReadandSendOneSecData(Cutterused_statistics, cip, ref allDataReadfromCIP, 36, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);

                                    bool[] Y6_temp = omronClients.ReadOneSecData(Y6, cip);
                                    Array.Copy(Y6_temp, allDataReadfromCIP.OEEInfo1Value, Y6_temp.Length);  //写到数据暂存区
                                    bool[] Manual_Andon_temp = omronClients.ReadOneSecData(Manual_Andon, cip);
                                    Array.Copy(Manual_Andon_temp, allDataReadfromCIP.OEEInfo2Value, Manual_Andon_temp.Length);  //写到数据暂存区

                                    var IECOEENumber = 20;
                                    var OEEValue = new bool[IECOEENumber];  //与IEC对应 20
                                    var OEEIndex = 0;

                                    //将Y6_temp和Manual_Andon 拼成一个OEE数组
                                    Array.Copy(Y6_temp, 0, OEEValue, OEEIndex, Y6_temp.Length);
                                    OEEIndex += Y6_temp.Length;
                                    Array.Copy(Manual_Andon_temp, 0, OEEValue, OEEIndex, Manual_Andon_temp.Length);

                                    

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

                                while (isThreadTwoRunning)
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
                                       // if (retIn.Content < 50 && retIn.Content >= 20)
                                            omronClients.SendSubArray(reya, ref allDataReadfromCIP, Station_Data, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);
                                    }
                                    else
                                    {
                                        logNet.WriteError("Auto_process[32] read failed");
                                        //Console.WriteLine("Auto_process[32] read failed");
                                    }

                                    retIn = cip.ReadInt32("Auto_process[18]");
                                    if (retIn.IsSuccess)
                                    {
                                        //if (retIn.Content < 50 && retIn.Content >= 25)
                                            omronClients.SendSubArray(dingfeng, ref allDataReadfromCIP, Station_Data, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);
                                    }
                                    else
                                    {
                                        logNet.WriteInfo("Auto_process[18] read failed ");
                                        //Console.WriteLine("Auto_process[18] read failed");
                                    }

                                    retIn = cip.ReadInt32("Auto_process[9]");
                                    if (retIn.IsSuccess)
                                    {
                                        //if (retIn.Content >= 40 && retIn.Content < 70)
                                            omronClients.SendSubArray(chongmo, ref allDataReadfromCIP, Station_Data, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);
                                    }
                                    else
                                    {
                                        logNet.WriteInfo("Auto_process[9] read failed ");
                                        //Console.WriteLine("Auto_process[9] read failed");
                                    }

                                    retIn = cip.ReadInt32("Auto_process[43]");
                                    if (retIn.IsSuccess)
                                    {
                                        //if (retIn.Content < 45 && retIn.Content >= 30)
                                            omronClients.SendSubArray(zuojiaofeng, ref allDataReadfromCIP, Station_Data, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);
                                    }
                                    else
                                    {
                                        logNet.WriteInfo("Auto_process[43] read failed ");
                                        //Console.WriteLine("Auto_process[43] read failed");
                                    }

                                    retIn = cip.ReadInt32("Auto_process[44]");
                                    if (retIn.IsSuccess)
                                    {
                                        //if (retIn.Content < 45 && retIn.Content >= 30)
                                            omronClients.SendSubArray(youjiaofeng, ref allDataReadfromCIP, Station_Data, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);
                                    }
                                    else
                                    {
                                        //logNet.WriteInfo("Auto_process[44] read failed ");
                                        Console.WriteLine("Auto_process[44] read failed");
                                    }


                                    retIn = cip.ReadInt32("Auto_process[47]");
                                    if (retIn.IsSuccess)
                                    {
                                        //if (retIn.Content < 60 && retIn.Content >= 30)
                                            omronClients.SendSubArray(cefeng, ref allDataReadfromCIP, Station_Data, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);
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
                                int NumberOfStation = 74;      //取1-74号工位

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

                                while (isThreadThreeRunning)
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
                                    allDataReadfromCIP.BarCode = sbBarCode;
                                    omronClients.ReadDeviceInfoConSturct1(EarCode, _cip[6], sbEarCode);
                                    allDataReadfromCIP.EarCode = sbEarCode;




                                    //整合到74个string中 舍弃下标0 并行发送数据                         
                                    for (int i = 0; i < NumberOfStation; i++)
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

                                        tempstring = combinedString.ToString();
                                        sendStringtoIEC[i].StringValue = tempstring; //整合到结构体数组中 （一把子把74个工位的数据发送给IEC）

                                        #region Grpc发送数据给IEC

                                        if (i == 73)
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

                            //将读取的值写入Excel 
                            thr[3] = new Thread(() =>
                            {                                                               
                                try
                                {                                
                                    var ExcelPath = "/opt/plcnext/apps/HGFZData.xlsx";
                                    //var ExcelPath = Directory.GetCurrentDirectory() + "\\HGFZData.xlsx";                        
                                    readExcel.setExcelCellValue(ExcelPath, "设备信息", "电芯条码地址采集值", allDataReadfromCIP.BarCode);
                                    readExcel.setExcelCellValue(ExcelPath, "设备信息", "极耳码地址采集值", allDataReadfromCIP.EarCode);
                                    readExcel.setExcelCellValue(ExcelPath, "加工工位（冲膜）", "采集值", allDataReadfromCIP.ChongMoValue);
                                    readExcel.setExcelCellValue(ExcelPath, "加工工位（热压）", "采集值", allDataReadfromCIP.ReYaValue);
                                    readExcel.setExcelCellValue(ExcelPath, "加工工位（顶封）", "采集值", allDataReadfromCIP.DingFengValue);
                                    readExcel.setExcelCellValue(ExcelPath, "加工工位（左角封）", "采集值", allDataReadfromCIP.ZuoJiaoFengValue);
                                    readExcel.setExcelCellValue(ExcelPath, "加工工位（右角封）", "采集值", allDataReadfromCIP.YouJiaoFengValue);
                                    readExcel.setExcelCellValue(ExcelPath, "加工工位（侧封）", "采集值", allDataReadfromCIP.CeFengValue);
                                    readExcel.setExcelCellValue(ExcelPath, "生产统计", "采集值", allDataReadfromCIP.ProductionDataValue);
                                    readExcel.setExcelCellValue(ExcelPath, "寿命管理", "采集值", allDataReadfromCIP.LifeManagementValue);
                                    readExcel.setExcelCellValue(ExcelPath, "OEE", "采集值", allDataReadfromCIP.OEEInfo1Value);
                                    readExcel.setExcelCellValue(ExcelPath, "OEE(2)", "采集值", allDataReadfromCIP.OEEInfo2Value);

                                    var tempFlag_finishWriteExcelFile = true;

                                    listWriteItem.Clear();
                                    listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["flag_finishWriteExcelFile"], Arp.Type.Grpc.CoreType.CtBoolean, tempFlag_finishWriteExcelFile));
                                    if (grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, grpcToolInstance.ServiceWriteRequestAddDatas(listWriteItem.ToArray()), new IDataAccessServiceWriteResponse(), options1))
                                    {
                                        Console.WriteLine("{0}      flag_finishWriteExcelFile写入IEC: success", DateTime.Now);
                                    }
                                    else
                                    {
                                        Console.WriteLine("{0}      flag_finishWriteExcelFile写入IEC: fail", DateTime.Now);
                                    }

                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine("Write data to Excel failed : {0} ", e);
                                }                                                                    
                                
                            });

                            #endregion
                            stepNumber = 100;

                        }
                        break;

                    

                    case 100:
                        {
                            //开启线程
                            if (thr[0].ThreadState == ThreadState.Unstarted && thr[1].ThreadState == ThreadState.Unstarted && thr[2].ThreadState == ThreadState.Unstarted)
                            {
                                try
                                {
                                    #region 开启三大数采线程

                                    isThreadOneRunning = true;
                                    thr[0].Start();  // 读1000ms数据

                                    isThreadTwoRunning = true;
                                    thr[1].Start(); //读六大工位信息

                                    isThreadThreeRunning = true;
                                    thr[2].Start();  //读设备信息
                                    

                                    //thr[3].Start();  //读设备信息
                                    #endregion

                                    //APP Status ： running
                                    listWriteItem.Clear();
                                    listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["AppStatus"], Arp.Type.Grpc.CoreType.CtInt32, 1));
                                    if (grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, grpcToolInstance.ServiceWriteRequestAddDatas(listWriteItem.ToArray()), new IDataAccessServiceWriteResponse(), options1))
                                    {
                                        Console.WriteLine("{0}      AppStatus写入IEC: success", DateTime.Now);
                                    }
                                    else
                                    {
                                        Console.WriteLine("{0}      AppStatus写入IEC: fail", DateTime.Now);
                                    }

                                }
                                catch
                                {
                                    Console.WriteLine("Thread quit");

                                    //APP Status ： Error
                                    listWriteItem.Clear();
                                    listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["AppStatus"], Arp.Type.Grpc.CoreType.CtInt32, -1));
                                    if (grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, grpcToolInstance.ServiceWriteRequestAddDatas(listWriteItem.ToArray()), new IDataAccessServiceWriteResponse(), options1))
                                    {
                                        Console.WriteLine("{0}      AppStatus写入IEC: success", DateTime.Now);
                                    }
                                    else
                                    {
                                        Console.WriteLine("{0}      AppStatus写入IEC: fail", DateTime.Now);
                                    }

                                    stepNumber = 1000;
                                    break;

                                }
                            }


                            #region IEC发送触发信号，重新读取Excel

                            dataAccessServiceReadSingleRequest = new IDataAccessServiceReadSingleRequest();
                            dataAccessServiceReadSingleRequest.PortName = nodeidDictionary["Switch_ReadExcelFile"];
                            if (grpcToolInstance.ReadSingleDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceReadSingleRequest, new IDataAccessServiceReadSingleResponse(), options1).BoolValue)
                            {
                                //复位信号点:Switch_WriteExcelFile                               
                                listWriteItem.Clear();
                                listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["Switch_ReadExcelFile"], Arp.Type.Grpc.CoreType.CtBoolean, false)); //Write Data to DataAccessService                                 
                                if (grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, grpcToolInstance.ServiceWriteRequestAddDatas(listWriteItem.ToArray()), new IDataAccessServiceWriteResponse(), options1))
                                {
                                    Console.WriteLine("{0}      Switch_ReadExcelFile写入IEC: success", DateTime.Now);
                                    logNet.WriteInfo(DateTime.Now.ToString() + "Switch_ReadExcelFile写入IEC: success");
                                }
                                else
                                {
                                    Console.WriteLine("{0}      Switch_ReadExcelFile写入IEC: fail", DateTime.Now);
                                    logNet.WriteError(DateTime.Now.ToString() + "Switch_ReadExcelFile写入IEC: fail");
                                }


                                //停止线程
                                isThreadOneRunning = false;
                                isThreadTwoRunning = false;
                                isThreadThreeRunning = false;

                                for (int i = 0; i < clientNum; i++)
                                {
                                    _cip[i].ConnectClose();
                                    Console.WriteLine(" CIP {0} Connect closed", i);
                                }

                                Thread.Sleep(1000);//等待线程退出

                                stepNumber = 10;
                            }

                            #endregion

                            #region 检测PLCnext和Omron PLC之间的连接
                            IPStatus iPStatus;
                            iPStatus = _cip[0].IpAddressPing();  //判断与PLC的物理连接状态
                            
                            if (iPStatus != 0)
                            {
                                //Console.WriteLine("Ping Omron PLC failed");
                                logNet.WriteError("Ping Omron PLC failed");

                            };
                            #endregion


                            #region IEC发送触发信号,将采集值写入Excel

                            dataAccessServiceReadSingleRequest = new IDataAccessServiceReadSingleRequest();
                            dataAccessServiceReadSingleRequest.PortName = nodeidDictionary["Switch_WriteExcelFile"];
                            if (grpcToolInstance.ReadSingleDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceReadSingleRequest, new IDataAccessServiceReadSingleResponse(), options1).BoolValue)
                            {
                                //复位信号点: Switch_WriteExcelFile
                                listWriteItem.Clear();
                                listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["Switch_WriteExcelFile"], Arp.Type.Grpc.CoreType.CtBoolean, false)); //Write Data to DataAccessService                                 
                                if (grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, grpcToolInstance.ServiceWriteRequestAddDatas(listWriteItem.ToArray()), new IDataAccessServiceWriteResponse(), options1))
                                {
                                    Console.WriteLine("{0}      Switch_WriteExcelFile: success", DateTime.Now);
                                    logNet.WriteInfo(DateTime.Now.ToString() + "Switch_WriteExcelFile: success");
                                }
                                else
                                {
                                    Console.WriteLine("{0}      Switch_WriteExcelFile: fail", DateTime.Now);
                                    logNet.WriteError(DateTime.Now.ToString() + "Switch_WriteExcelFile: fail");
                                }

                                try
                                {
                                    logNet.WriteInfo("-----" + thr[3].ThreadState.ToString());

                                    thr[3].Start();

                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine("Write data to Excel failed : {0} ", e);
                                }

                            }

                            #endregion

                            Thread.Sleep(1000);

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