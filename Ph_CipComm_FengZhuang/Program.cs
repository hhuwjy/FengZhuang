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
        
        public static ILogNet logNet = new LogNetFileSize(logsFile, 5 * 1024 * 1024); //限制了日志大小

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
        static int clientNum = 3; //CIP的Client最大32，开启3个CIP Client    
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

        static DeviceInfoConSturct_CIP[] Auto_Process;
        static DeviceInfoConSturct_CIP[] Clear_Manual;

        //设备信息里的离散结构体数据和数组数据
        static DeviceInfoConSturct_CIP[] Battery_Memory;
        static DeviceInfoConSturct_CIP[] BarCode;
        static DeviceInfoConSturct_CIP[] EarCode;

        //九大工位数据
        static StationInfoStruct_CIP[] chongmo;
        static StationInfoStruct_CIP[] reya_1;
        static StationInfoStruct_CIP[] reya_2;
        static StationInfoStruct_CIP[] reya_3;
        static StationInfoStruct_CIP[] reya_4;

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
        

        ////（六大工位，生产统计，寿命管理 ，电芯码 、极耳码、OEE数据）数据暂存结构体
        //static AllDataReadfromCIP allDataReadfromCIP = new AllDataReadfromCIP();

        // 设备总览
        static DeviceInfoStruct_IEC[] deviceInfoStruct_IEC;

        #endregion

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
            int IecTriggersNumber = 0;

            AllDataReadfromCIP allDataReadfromCIP = new AllDataReadfromCIP();

          

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

                            stepNumber = 6;
                        }
                        break;


                    case 6:
                        {

                            #region 从xml获取nodeid，Grpc发送到对应变量时使用，注意xml中的别名要和对应类的属性名一致 
                            try
                            {
                                const string filePath = "/opt/plcnext/apps/GrpcSubscribeNodes.xml";             //EPC中存放的路径  
                                //const string filePath = "D:\\2024\\Work\\12-冠宇数采项目\\ReadFromStructArray\\FengZhuang_EIP\\Ph_CipComm_FengZhuang\\GrpcSubscribeNodes\\GrpcSubscribeNodes.xml";  //PC中存放的路径 

                                nodeidDictionary = grpcToolInstance.getNodeIdDictionary(filePath);  //将xml中的值写入字典中
                                logNet.WriteInfo("NodeID xml文件读取成功");

                            }
                            catch (Exception e)
                            {
                                logNet.WriteError("NodeID xml文件读取失败：" + e);
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

                            logNet.WriteInfo("封装设备数采APP已启动");

                            //string excelFilePath = Directory.GetCurrentDirectory() + "\\HGFZData.xlsx";     //PC端测试路径
                            string excelFilePath = "/opt/plcnext/apps/HGFZData.xlsx";                         //EPC存放路径

                            XSSFWorkbook excelWorkbook = readExcel.connectExcel(excelFilePath);

                            Console.WriteLine("ExcelWorkbook read {0}", excelWorkbook != null ? "success" : "fail");
                            if (excelWorkbook != null)
                            {
                                logNet.WriteInfo("Excel读取成功");
                            }
                            else
                            {
                                logNet.WriteError("Excel读取失败");
                            }



                            // 给IEC发送 Excel读取成功的信号
                            var tempFlag_finishReadExcelFile = true;

                            listWriteItem.Clear();
                            listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["flag_finishReadExcelFile"], Arp.Type.Grpc.CoreType.CtBoolean, tempFlag_finishReadExcelFile));
                            if (grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, grpcToolInstance.ServiceWriteRequestAddDatas(listWriteItem.ToArray()), new IDataAccessServiceWriteResponse(), options1))
                            {
                                //Console.WriteLine("{0}      flag_finishReadExcelFile写入IEC: success", DateTime.Now);
                                logNet.WriteInfo("[Grpc]", "flag_finishReadExcelFile 写入IEC成功");
                            }
                            else
                            {
                                //Console.WriteLine("{0}      flag_finishReadExcelFile写入IEC: fail", DateTime.Now);
                                logNet.WriteError("[Grpc]", "flag_finishReadExcelFile 写入IEC失败");
                            }
                            #endregion


                            #region 将readExcel变量中的值，写入对应的实例化结构体中

                            // 九大工位（100ms)
                            chongmo = readExcel.ReadStationInfo_Excel(excelWorkbook, "加工工位（冲膜）");
                            reya_1 = readExcel.ReadStationInfo_Excel(excelWorkbook, "加工工位（热压1）");
                            reya_2 = readExcel.ReadStationInfo_Excel(excelWorkbook, "加工工位（热压2）");
                            reya_3 = readExcel.ReadStationInfo_Excel(excelWorkbook, "加工工位（热压3）");
                            reya_4 = readExcel.ReadStationInfo_Excel(excelWorkbook, "加工工位（热压4）");
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
                            Auto_Process = readExcel.ReadOneDeviceInfoConSturct1Info_Excel(excelWorkbook, "设备信息", "工位加工中(DINT)");
                            Clear_Manual = readExcel.ReadOneDeviceInfoConSturct1Info_Excel(excelWorkbook, "设备信息", "电芯记忆清除按钮(BOOL)");
                            Battery_Memory = readExcel.ReadOneDeviceInfoConSturct1Info_Excel(excelWorkbook, "设备信息", "电芯记忆(BOOL)");
                            BarCode = readExcel.ReadOneDeviceInfoConSturct2Info_Excel(excelWorkbook, "设备信息", "电芯条码地址(REAL)");
                            EarCode = readExcel.ReadOneDeviceInfoConSturct2Info_Excel(excelWorkbook, "设备信息", "极耳码地址(REAL)");

                            #endregion


                            #region 单独设备总览表 + 发送

                            var deviceInfoStructList_IEC = new DeviceInfoStructList_IEC();

                            deviceInfoStruct_IEC = readExcel.ReadDeviceInfo_Excel(excelWorkbook, "封装设备总览");


                            deviceInfoStructList_IEC.iCount = (short)deviceInfoStruct_IEC.Length;
                            for (int i = 0; i<deviceInfoStruct_IEC.Length; i++)
                            {
                                deviceInfoStructList_IEC.arrDeviceInfo[i].strDeviceName = deviceInfoStruct_IEC[i].strDeviceName;
                                deviceInfoStructList_IEC.arrDeviceInfo[i].strDeviceCode = deviceInfoStruct_IEC[i].strDeviceCode;
                                deviceInfoStructList_IEC.arrDeviceInfo[i].iStationCount = deviceInfoStruct_IEC[i].iStationCount;
                                deviceInfoStructList_IEC.arrDeviceInfo[i].strPLCType = deviceInfoStruct_IEC[i].strPLCType;
                                deviceInfoStructList_IEC.arrDeviceInfo[i].strProtocol = deviceInfoStruct_IEC[i].strProtocol;
                                deviceInfoStructList_IEC.arrDeviceInfo[i].strIPAddress = deviceInfoStruct_IEC[i].strIPAddress;
                                deviceInfoStructList_IEC.arrDeviceInfo[i].iPort = deviceInfoStruct_IEC[i].iPort;
                            }

                            listWriteItem.Clear();
                            try
                            {
                                listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["DeviceInfo"], Arp.Type.Grpc.CoreType.CtStruct, deviceInfoStructList_IEC));
                                var writeItemsArray = listWriteItem.ToArray();
                                var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                                bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                            }
                            catch (Exception e)
                            {
                                //Console.WriteLine("ERRO: {0}，{1}", e, nodeidDictionary.GetValueOrDefault("DeviceInfo"));
                                logNet.WriteError("[Grpc]", "封装设备总览表格 数据发送 失败" + e);
                            }

                            #endregion

                            stepNumber = 15;
                            


                        }

                        break;

                    case 15:
                        {
                           
                            #region 读取并发送1000ms数据的点位名

                            //实例化发给IEC的 1000ms数据的点位名 结构体
                            var OneSecNameStruct = new OneSecPointNameStruct_IEC();

                            // 功能安全、生产统计和寿命管理 的点位名
                            omronClients.ReadPointName(Sys_Manual, ref OneSecNameStruct);
                            omronClients.ReadPointName(Production_statistics, ref OneSecNameStruct);
                            omronClients.ReadPointName(Cutterused_statistics, ref OneSecNameStruct);

                            //OEE 的点位名
                            omronClients.ReadPointName(Y6, Manual_Andon, ref OneSecNameStruct);

                            //报警信息的点位名
                            var stringnumber = Vacuum_Alarm.Length + Senor_Alarm.Length + Motor_POTLimit_Err.Length + Motor_NOTLimit_Err.Length
                                           + Motor_Prevent_Promt.Length + Exception_information.Length + Temperature_Alarm.Length
                                           + Cylinder_Reset_Promt.Length + Motor_Reset_Promt.Length + Stopstate_Error.Length + Grating_Error.Length
                                           + Scapegoat_Error.Length + Door_Error.Length + out_power.Length;

                            List<OneSecInfoStruct_CIP[]> alarmGroups = new List<OneSecInfoStruct_CIP[]> {
                                Vacuum_Alarm, Senor_Alarm, Motor_POTLimit_Err, Motor_NOTLimit_Err, Motor_Prevent_Promt,
                                Exception_information,  Temperature_Alarm, Cylinder_Reset_Promt,
                                Motor_Reset_Promt, Stopstate_Error, Grating_Error, Scapegoat_Error, Door_Error, out_power
                            };

                            omronClients.ReadPointName(alarmGroups, stringnumber, ref OneSecNameStruct);

                            //Grpc发送1000ms数据点位名结构体
                            listWriteItem.Clear();
                            try
                            {
                                listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary.GetValueOrDefault("OneSecNameStruct"), Arp.Type.Grpc.CoreType.CtStruct, OneSecNameStruct));
                                var writeItemsArray = listWriteItem.ToArray();
                                var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                                bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                            }
                            catch (Exception e)
                            {
                                logNet.WriteError("[Grpc]", " 1000ms数据的点位名发送失败：" + e);
                                //Console.WriteLine("ERRO: {0}", e);
                            }


                            #endregion


                            #region 读取并发送六大工位的点位名

                            var ProcessStationNameStruct = new ProcessStationNameStruct_IEC();

                            List<StationInfoStruct_CIP[]> StationDataStruct = new List<StationInfoStruct_CIP[]>
                            { chongmo, reya_1, reya_2, reya_3, reya_4, dingfeng, zuojiaofeng, youjiaofeng, cefeng };

                            omronClients.ReadPointName(StationDataStruct, ref ProcessStationNameStruct);

                            //Grpc发送1000ms数据点位名结构体
                            listWriteItem.Clear();
                            try
                            {
                                listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary.GetValueOrDefault("ProcessStationNameStruct"), Arp.Type.Grpc.CoreType.CtStruct, ProcessStationNameStruct));
                                var writeItemsArray = listWriteItem.ToArray();
                                var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                                bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                            }
                            catch (Exception e)
                            {
                                logNet.WriteError("[Grpc]", " 加工工位的点位名发送失败：" + e);
                                //Console.WriteLine("ERRO: {0}", e);
                            }

                            #endregion


                            #region 读取并发送 设备信息表里的信息 （后工位序号 + 生成虚拟码）

                            omronClients.ReadandSendStaionInfo(Auto_Process, grpcToolInstance, nodeidDictionary, grpcDataAccessServiceClient, options1);

                            #endregion

                            stepNumber = 20;

                           
                        }


                        break;



                    case 20:
                        {

                            #region CIP连接   

                            logNet.WriteInfo("[CIP]", "CIP连接设备的ip地址为：" + deviceInfoStruct_IEC[0].strIPAddress);

                            for (int i = 0; i < clientNum; i++)
                            {
                                _cip[i] = new OmronConnectedCipNet(deviceInfoStruct_IEC[0].strIPAddress);  //  填写欧姆龙PLC的IP地址
                                var retConnect = _cip[i].ConnectServer();

                                //Console.WriteLine("num {0} connect: {1}!", i, retConnect.IsSuccess ? "success" : "fail");
               
                                if (retConnect.IsSuccess)
                                {
                                    logNet.WriteInfo("[CIP]","CIP连接成功" + i.ToString());                                    
                                }
                                else
                                {
                                    logNet.WriteError("[CIP]","CIP连接失败" + i.ToString());

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
                                int lastIndex = out_power[out_power.Length - 1].varIndex;
                                bool[] Alarm_Data = new bool[lastIndex + 1];

                                var DeviceDataStruct = new DeviceDataStruct_IEC();

                                while (isThreadOneRunning)
                                {
                              
                                    TimeSpan start = new TimeSpan(DateTime.Now.Ticks);

                                    #region 读取报警信号 （读总数组，进行拼接）
                                   
                                    //读取Upload_Err总数组

                                    OperateResult<bool[]> retbool = cip.ReadBool(Vacuum_Alarm[0].varName, (ushort)(lastIndex + 1));
                                    if (retbool.IsSuccess)
                                    {
                                        Alarm_Data = retbool.Content;
                                    }
                                    else
                                    {
                                        //Console.WriteLine(Vacuum_Alarm[0].varName + "Array Read Failed");
                                        logNet.WriteError("[CIP]","报警数据读取失败");
                                    }

                                    //发送给IEC的
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
                                        Array.Copy(Alarm_Data, StartIndex, DeviceDataStruct.Value_ALM, ArrayIndex, alarmGroup.Length);
                                        ArrayIndex += alarmGroup.Length;
                                    }                                        
                                        
                                    #endregion


                                    #region 读取非报警信号

                                    //读取非报警信号 
                                    omronClients.ReadOneSecData(Sys_Manual, cip, ref  allDataReadfromCIP, ref DeviceDataStruct);
                                    omronClients.ReadOneSecData(Production_statistics, cip, ref allDataReadfromCIP, ref DeviceDataStruct);
                                    omronClients.ReadOneSecData(Cutterused_statistics, cip, ref allDataReadfromCIP, ref DeviceDataStruct);


                                    bool[] Y6_temp = omronClients.ReadOneSecData(Y6, cip);
                                    Array.Copy(Y6_temp, allDataReadfromCIP.OEEInfo1Value, Y6_temp.Length);  //写到数据暂存区（写入Excel）

                                    bool[] Manual_Andon_temp = omronClients.ReadOneSecData(Manual_Andon, cip);
                                    Array.Copy(Manual_Andon_temp, allDataReadfromCIP.OEEInfo2Value, Manual_Andon_temp.Length);  //写到数据暂存区（写入Excel）


                                    var OEEIndex = 0;

                                    //将Y6_temp和Manual_Andon 拼成一个OEE数组 写入 DeviceDataStruct （发给IEC）结构体中
                                    Array.Copy(Y6_temp, 0, DeviceDataStruct.Value_OEE, OEEIndex, Y6_temp.Length);
                                    OEEIndex += Y6_temp.Length;
                                    Array.Copy(Manual_Andon_temp, 0, DeviceDataStruct.Value_OEE, OEEIndex, Manual_Andon_temp.Length);

                                    #endregion


                                    //Grpc 发送1000ms数据采集值

                                    listWriteItem.Clear();
                                  
                                    try
                                    {
                                        listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["OneSecDataValue"], Arp.Type.Grpc.CoreType.CtStruct, DeviceDataStruct));
                                        var writeItemsArray = listWriteItem.ToArray();
                                        var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                                        bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                                    }
                                    catch (Exception e)
                                    {
                                        //logNet.WriteError("[Grpc]", "OEE数据发送失败：" + e);
                                        Console.WriteLine("ERRO: {0}", e, nodeidDictionary.GetValueOrDefault("OneSecDataValue"));
                                    }

   
                                    //计算从开始读到读完的时间
                                    TimeSpan end = new TimeSpan(DateTime.Now.Ticks);
                                    DateTime nowDisplay = DateTime.Now;
                                    TimeSpan dur = (end - start).Duration();

                                    if (dur.TotalMilliseconds < 1000)
                                    {
                                        int sleepTime = 1000 - (int)dur.TotalMilliseconds;
                                        //Console.WriteLine("Thread ReadOnceSecInfo read time:{0} read Duration:{1}", nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff"), dur.TotalMilliseconds);
                                        Thread.Sleep(sleepTime);
                                    }
                                    else
                                    {
                                        logNet.WriteInfo("Thread ReadOnceSecInfo Read Time : " + (dur.TotalMilliseconds).ToString());
                                        //Console.WriteLine("Thread ReadOnceSecInfo read time:{0} read Duration:{1}", nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff"), dur.TotalMilliseconds);

                                    }

                                }
                            });

                            //读九大工位的数据
                            thr[1] = new Thread(() =>
                            {
                                var cip = _cip[1];
                                var ProcessStationDataValue = new UDT_ProcessStationDataValue();
                                var listWriteItem = new List<WriteItem>();

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
                                        //Console.WriteLine(Vacuum_Alarm[0].varName + "Array Read Failed");
                                        logNet.WriteError("[CIP]", Vacuum_Alarm[0].varName + "数据读取失败");
                                    }
                                    ProcessStationDataValue.iDataCount = 9;

                                    // 将大数组里的数组按照工位写进 ProcessStationDataValue 结构体中
                                    omronClients.WriteSubArray(chongmo, ref allDataReadfromCIP, Station_Data, ref ProcessStationDataValue);                                                                      
                                    omronClients.WriteSubArray(reya_1, ref allDataReadfromCIP, Station_Data, ref ProcessStationDataValue);
                                    omronClients.WriteSubArray(reya_2, ref allDataReadfromCIP, Station_Data, ref ProcessStationDataValue);
                                    omronClients.WriteSubArray(reya_3, ref allDataReadfromCIP, Station_Data, ref ProcessStationDataValue);
                                    omronClients.WriteSubArray(reya_4, ref allDataReadfromCIP, Station_Data, ref ProcessStationDataValue);                             
                                    omronClients.WriteSubArray(dingfeng, ref allDataReadfromCIP, Station_Data, ref ProcessStationDataValue);
                                    omronClients.WriteSubArray(zuojiaofeng, ref allDataReadfromCIP, Station_Data, ref ProcessStationDataValue);
                                    omronClients.WriteSubArray(youjiaofeng, ref allDataReadfromCIP, Station_Data, ref ProcessStationDataValue);
                                    omronClients.WriteSubArray(cefeng, ref allDataReadfromCIP, Station_Data, ref ProcessStationDataValue);


                                    //Grpc 发送加工工位数据采集值

                                    listWriteItem.Clear();

                                    try
                                    {
                                        listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["ProcessStationData"], Arp.Type.Grpc.CoreType.CtStruct, ProcessStationDataValue));
                                        var writeItemsArray = listWriteItem.ToArray();
                                        var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                                        bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("[Grpc]", "加工工位数据发送失败：" + e);
                                 
                                    }

                                    #endregion

                                    TimeSpan end = new TimeSpan(DateTime.Now.Ticks);
                                    DateTime nowDisplay = DateTime.Now;
                                    TimeSpan dur = (end - start).Duration();

                                   
                                    if (dur.TotalMilliseconds < 100)
                                    {
                                        int sleepTime = 100 - (int)dur.TotalMilliseconds;
                                        Thread.Sleep(sleepTime);
                                    }
                                    else
                                    {
                                        logNet.WriteInfo("Thread ProcessStation Data Read Time : " + (dur.TotalMilliseconds).ToString());
                                        //Console.WriteLine("Thread ReadStationInfo1 read time:{0} read Duration:{1}", nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff"), dur.TotalMilliseconds);

                                    }

                                }
                            });

                            //读74个工位的数据
                            thr[2] = new Thread(() =>
                            {

                                var StationListlnfo = new UDT_StationListlnfo();
                                StationListlnfo.iDataCount = (short)Auto_Process.Length;
                                var cip = _cip[2];

                                var listWriteItem = new List<WriteItem>();
                                WriteItem[] writeItems = new WriteItem[] { };

                                while (isThreadThreeRunning)
                                {
                                  
                                    TimeSpan start = new TimeSpan(DateTime.Now.Ticks);

                                    
                                    omronClients.ReadDeviceInfoConSturct(Auto_Process, cip, ref allDataReadfromCIP, ref StationListlnfo);
                                    omronClients.ReadDeviceInfoConSturct(Clear_Manual, cip, ref allDataReadfromCIP, ref StationListlnfo);
                                    omronClients.ReadDeviceInfoConSturct(Battery_Memory, cip, ref allDataReadfromCIP, ref StationListlnfo);
                                    omronClients.ReadDeviceInfoConSturct(BarCode, cip, ref allDataReadfromCIP, ref StationListlnfo);
                                    omronClients.ReadDeviceInfoConSturct(EarCode, cip, ref allDataReadfromCIP, ref StationListlnfo);

                                 
                                    // Grpc发送数据给IEC                     
                                    try
                                    {
                                        listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["StationListlnfo"], Arp.Type.Grpc.CoreType.CtStruct, StationListlnfo));
                                        var writeItemsArray = listWriteItem.ToArray();
                                        var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                                        bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("[Grpc]", "设备信息数据发送失败：" + e);
                                        //Console.WriteLine("ERRO: {0}，{1}", e, nodeidDictionary.GetValueOrDefault(i.ToString()));
                                    }
                                    listWriteItem.Clear();                               


                                    TimeSpan end = new TimeSpan(DateTime.Now.Ticks);
                                    DateTime nowDisplay = DateTime.Now;
                                    TimeSpan dur = (end - start).Duration();

                                    

                                    if (dur.TotalMilliseconds < 100)
                                    {
                                        int sleepTime = 100 - (int)dur.TotalMilliseconds;
                                        Thread.Sleep(sleepTime);
                                    }
                                    else
                                    {
                                        logNet.WriteInfo("Thread ReadDeviceInfo read time : " + (dur.TotalMilliseconds).ToString());
                                        //Console.WriteLine("Thread ReadDeviceInfo read time:{0} read Duration:{1}", nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff"), dur.TotalMilliseconds);

                                    }

                                }
                            });

                            #endregion

                            stepNumber = 100;

                        }
                        break;

                    

                    case 100:
                        {

                            #region 开启三大数采线程
                            if (thr[0].ThreadState == ThreadState.Unstarted && thr[1].ThreadState == ThreadState.Unstarted && thr[2].ThreadState == ThreadState.Unstarted)
                            {
                                try
                                {
             
                                    isThreadOneRunning = true;
                                    thr[0].Start();  // 读1000ms数据

                                    isThreadTwoRunning = true;
                                    thr[1].Start(); //读九个加工工位信息

                                    isThreadThreeRunning = true;
                                    thr[2].Start();  //读设备信息



                                    //APP Status ： running
                                    listWriteItem.Clear();
                                    listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["AppStatus"], Arp.Type.Grpc.CoreType.CtInt32, 1));
                                    if (grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, grpcToolInstance.ServiceWriteRequestAddDatas(listWriteItem.ToArray()), new IDataAccessServiceWriteResponse(), options1))
                                    {
                                        logNet.WriteInfo("[Grpc]", "AppStatus 写入IEC成功");
                                        //Console.WriteLine("{0}      AppStatus写入IEC: success", DateTime.Now);
                                    }
                                    else
                                    {
                                        //Console.WriteLine("{0}      AppStatus写入IEC: fail", DateTime.Now);
                                        logNet.WriteError("[Grpc]", "AppStatus 写入IEC失败");
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
                                        logNet.WriteInfo("[Grpc]", "AppStatus 写入IEC成功");
                                        //Console.WriteLine("{0}      AppStatus写入IEC: success", DateTime.Now);
                                    }
                                    else
                                    {
                                        //Console.WriteLine("{0}      AppStatus写入IEC: fail", DateTime.Now);
                                        logNet.WriteError("[Grpc]", "AppStatus 写入IEC失败");
                                    }

                                    stepNumber = 1000;
                                    break;

                                }
                            }
                            #endregion


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
                                    //Console.WriteLine("{0}      Switch_ReadExcelFile写入IEC: success", DateTime.Now);
                                    logNet.WriteInfo("[Grpc]", "Switch_ReadExcelFile 写入IEC成功");
                                }
                                else
                                {
                                    //Console.WriteLine("{0}      Switch_ReadExcelFile写入IEC: fail", DateTime.Now);
                                    logNet.WriteError("[Grpc]", "Switch_ReadExcelFile 写入IEC失败");
                                }


                                //停止线程
                                isThreadOneRunning = false;
                                isThreadTwoRunning = false;
                                isThreadThreeRunning = false;

                                for (int i = 0; i < clientNum; i++)
                                {
                                    _cip[i].ConnectClose();
                                    //Console.WriteLine(" CIP {0} Connect closed", i);
                                    logNet.WriteInfo("[CIP]", "CIP连接断开" + i.ToString());
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
                                logNet.WriteError("[CIP]", "Ping Omron PLC failed");

                                //APP Status ： Error
                                listWriteItem.Clear();
                                listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["AppStatus"], Arp.Type.Grpc.CoreType.CtInt32, -2));
                                if (grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, grpcToolInstance.ServiceWriteRequestAddDatas(listWriteItem.ToArray()), new IDataAccessServiceWriteResponse(), options1))
                                {
                                    logNet.WriteInfo("[Grpc]", "AppStatus 写入IEC成功");
                                    //Console.WriteLine("{0}      AppStatus写入IEC: success", DateTime.Now);
                                }
                                else
                                {
                                    //Console.WriteLine("{0}      AppStatus写入IEC: fail", DateTime.Now);
                                    logNet.WriteError("[Grpc]", "AppStatus 写入IEC失败");
                                }


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
                                    //Console.WriteLine("{0}      Switch_WriteExcelFile: success", DateTime.Now);
                                    logNet.WriteInfo("[Grpc]", "Switch_WriteExcelFile 写入IEC成功");
                                }
                                else
                                {
                                    //Console.WriteLine("{0}      Switch_WriteExcelFile: fail", DateTime.Now);
                                    logNet.WriteError("[Grpc]", "Switch_WriteExcelFile 写入IEC失败");
                                }

                                //将读取的值写入Excel 
                                thr[3] = new Thread(() =>
                                {
                                    var ExcelPath = "/opt/plcnext/apps/HGFZData.xlsx";
                                    //var ExcelPath = Directory.GetCurrentDirectory() + "\\HGFZData.xlsx";

                                    var allDataReadfromCIP_temp = allDataReadfromCIP;  //将数据缓存区的值赋给临时变量

                                    #region 将数据缓存区的值写入Excel

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath, "设备信息", "电芯条码地址采集值", allDataReadfromCIP_temp.BarCode);
                                        logNet.WriteInfo("WriteData", "电芯条码地址采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("WriteData", "电芯条码地址采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath, "加工工位（冲膜）", "采集值", allDataReadfromCIP_temp.ChongMoValue);
                                        logNet.WriteInfo("WriteData", "加工工位（冲膜）采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        logNet.WriteError("WriteData", "电芯条码地址采集值写入Excel失败原因: " + e);
                                        //Console.WriteLine("加工工位（冲膜）采集值写入Excel失败原因: {0} ", e);

                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath, "设备信息", "极耳码地址采集值", allDataReadfromCIP_temp.EarCode);
                                        logNet.WriteInfo("WriteData", "极耳码地址采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        // Console.WriteLine("极耳码地址采集值写入Excel失败原因: {0} ", e);
                                        logNet.WriteError("WriteData", "极耳码地址采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath, "加工工位（热压1）", "采集值", allDataReadfromCIP_temp.ReYaValue_1);
                                        logNet.WriteInfo("加工工位（热压1）采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        //Console.WriteLine("加工工位（热压）采集值写入Excel失败原因: {0} ", e);
                                        logNet.WriteError("WriteData", "加工工位（热压1）采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath, "加工工位（热压2）", "采集值", allDataReadfromCIP_temp.ReYaValue_2);
                                        logNet.WriteInfo("加工工位（热压2）采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        //Console.WriteLine("加工工位（热压）采集值写入Excel失败原因: {0} ", e);
                                        logNet.WriteError("WriteData", "加工工位（热压2）采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath, "加工工位（热压3）", "采集值", allDataReadfromCIP_temp.ReYaValue_3);
                                        logNet.WriteInfo("加工工位（热压3）采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        //Console.WriteLine("加工工位（热压）采集值写入Excel失败原因: {0} ", e);
                                        logNet.WriteError("WriteData", "加工工位（热压3）采集值写入Excel失败原因: " + e);
                                    }


                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath, "加工工位（热压4）", "采集值", allDataReadfromCIP_temp.ReYaValue_4);
                                        logNet.WriteInfo("加工工位（热压4）采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        //Console.WriteLine("加工工位（热压）采集值写入Excel失败原因: {0} ", e);
                                        logNet.WriteError("WriteData", "加工工位（热压4）采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath, "加工工位（顶封）", "采集值", allDataReadfromCIP_temp.DingFengValue);
                                        logNet.WriteInfo("加工工位（顶封）采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        //Console.WriteLine("加工工位（顶封）采集值写入Excel失败原因: {0} ", e);
                                        logNet.WriteError("WriteData", "加工工位（顶封）采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath, "加工工位（左角封）", "采集值", allDataReadfromCIP_temp.ZuoJiaoFengValue);
                                        logNet.WriteInfo("加工工位（左角封）采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        //Console.WriteLine("加工工位（左角封）采集值写入Excel失败原因: {0} ", e);
                                        logNet.WriteError("WriteData", "加工工位（左角封）采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath, "加工工位（右角封）", "采集值", allDataReadfromCIP_temp.YouJiaoFengValue);
                                        logNet.WriteInfo("加工工位（右角封）采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        //Console.WriteLine("加工工位（右角封）采集值写入Excel失败原因: {0} ", e);
                                        logNet.WriteError("WriteData", "加工工位（右角封）采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath, "加工工位（侧封）", "采集值", allDataReadfromCIP_temp.CeFengValue);
                                        logNet.WriteInfo("加工工位（侧封）采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        //Console.WriteLine("加工工位（侧封）采集值写入Excel失败原因: {0} ", e);
                                        logNet.WriteError("WriteData", "加工工位（侧封）采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath, "生产统计", "采集值", allDataReadfromCIP_temp.ProductionDataValue);
                                        logNet.WriteInfo("生产统计采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        //Console.WriteLine("生产统计采集值写入Excel失败原因: {0} ", e);
                                        logNet.WriteError("WriteData", "生产统计采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath, "寿命管理", "采集值", allDataReadfromCIP_temp.LifeManagementValue);
                                        logNet.WriteInfo("寿命管理采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        //Console.WriteLine("寿命管理采集值写入Excel失败原因: {0} ", e);
                                        logNet.WriteError("WriteData", "寿命管理采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath, "功能开关", "采集值", allDataReadfromCIP_temp.FunctionEnableValue);
                                        logNet.WriteInfo("功能开关采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        //Console.WriteLine("功能开关采集值写入Excel失败原因: {0} ", e);
                                        logNet.WriteError("WriteData", "功能开关采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath, "OEE", "采集值", allDataReadfromCIP_temp.OEEInfo1Value);
                                        logNet.WriteInfo("OEE采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        //Console.WriteLine("OEE采集值写入Excel失败原因: {0} ", e);
                                        logNet.WriteError("WriteData", "OEE采集值写入Excel失败原因: " + e);
                                    }

                                    try
                                    {
                                        var result = readExcel.setExcelCellValue(ExcelPath, "OEE(2)", "采集值", allDataReadfromCIP_temp.OEEInfo2Value);
                                        logNet.WriteInfo("OEE(2)采集值写入Excel: " + (result ? "成功" : "失败"));
                                    }
                                    catch (Exception e)
                                    {
                                        //Console.WriteLine("OEE(2)采集值写入Excel失败原因: {0} ", e);
                                        logNet.WriteError("WriteData", "OEE(2)采集值写入Excel失败原因: " + e);
                                    }
                                    #endregion

                                    //给IEC写入 采集值写入成功的信号
                                    var tempFlag_finishWriteExcelFile = true;

                                    listWriteItem.Clear();
                                    listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary["flag_finishWriteExcelFile"], Arp.Type.Grpc.CoreType.CtBoolean, tempFlag_finishWriteExcelFile));
                                    if (grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, grpcToolInstance.ServiceWriteRequestAddDatas(listWriteItem.ToArray()), new IDataAccessServiceWriteResponse(), options1))
                                    {
                                        //Console.WriteLine("{0}      flag_finishWriteExcelFile写入IEC: success", DateTime.Now);
                                        logNet.WriteInfo("[Grpc]", "flag_finishWriteExcelFile 写入IEC成功");
                                    }
                                    else
                                    {
                                        //Console.WriteLine("{0}      flag_finishWriteExcelFile写入IEC: fail", DateTime.Now);
                                        logNet.WriteError("[Grpc]", "flag_finishWriteExcelFile 写入IEC失败");
                                    }

                                    IecTriggersNumber = 0;  //为了防止IEC连续两次赋值true

                                });

                                IecTriggersNumber++;

                                if (IecTriggersNumber == 1)
                                {
                                    thr[3].Start();
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