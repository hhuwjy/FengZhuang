using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HslCommunication;
using HslCommunication.Profinet.Omron;
using System.Threading;
using System.Security.Cryptography;
using System.Diagnostics.CodeAnalysis;
using System.Numerics;
using System.Runtime.CompilerServices;
using System.Collections;
using Grpc.Core;
using static Arp.Plc.Gds.Services.Grpc.IDataAccessService;
using Arp.Plc.Gds.Services.Grpc;
using Grpc.Net.Client;
using static Ph_CipComm_FengZhuang.GrpcTool;
using System.Net.Sockets;
using System.Drawing;
using Opc.Ua;
using NPOI.SS.Formula.Functions;
using HslCommunication.LogNet;
using Microsoft.Extensions.Logging;
using static Ph_CipComm_FengZhuang.UserStruct;
using static Ph_CipComm_FengZhuang.Program;
using NPOI.Util;




namespace Ph_CipComm_FengZhuang
{

    class OmronComm
    {

        #region Function 读取设备信息（以数组形式一起读上来，再按照序号写入对应的工位里）

        // 读 工位加工中 数据
        public void ReadDeviceInfoConSturct(DeviceInfoConSturct_CIP[] input, OmronConnectedCipNet cip, ref AllDataReadfromCIP allDataReadfromCIP, ref UDT_StationListlnfo StationListlnfo) 
        {
           string ReadObject = input[0].varName;   
                                                    
           switch (ReadObject)
           {
                case "Auto_process":
                    {
                        ushort Auto_Process_Length = 86;
                        OperateResult<int[]> ret = cip.ReadInt32(ReadObject, Auto_Process_Length);
                        if (ret.IsSuccess)
                        {
                            for (int i = 0; i < input.Length; i++)
                            {               
                                StationListlnfo.arrDataPoint[input[i].stationNumber - 1].diProcessData = ret.Content[input[i].varIndex];
                            }

                        }
                        else
                        {
                            logNet.WriteError("[CIP]", ReadObject + "读取失败");
                            //Console.WriteLine(ReadObject + "read failed");

                        }

                    }
                    break;

                case "Station_Cell":
                    {
                        ushort length = 57;        // 数组长度为硬编码，由Excel读出，不知后续需要是否需要更改
                        OperateResult<bool[]> ret = cip.ReadBool(ReadObject, length);
                        if (ret.IsSuccess)
                        {
                            for (int i = 0; i < input.Length; i++)
                            {                             
                                StationListlnfo.arrDataPoint[input[i].stationNumber - 1].xCellMem = ret.Content[input[i].varIndex];
                            }

                        }
                        else
                        {
                            logNet.WriteError("[CIP]", ReadObject + "读取失败");
                            //Console.WriteLine(ReadObject + "read failed");

                        }
                    }
                    break;

                case "Clear_Manual":
                    {
                        ushort length = 76; // 数组长度为硬编码，由Excel读出，不知后续需要是否需要更改
                        OperateResult<bool[]> ret = cip.ReadBool(ReadObject, length);
                        if (ret.IsSuccess)
                        {
                            for (int i = 0; i < input.Length; i++)
                            {                             
                                StationListlnfo.arrDataPoint[input[i].stationNumber - 1].xCellMemClear = ret.Content[input[i].varIndex];
                            }

                        }
                        else
                        {
                            logNet.WriteError("[CIP]", ReadObject + "读取失败");
                            //Console.WriteLine(ReadObject + "read failed");
                        }
                    }
                    break;

                case "Station_BarCode":
                    {
                        ushort length = 500; // 数组长度为硬编码，由Excel读出，不知后续需要是否需要更改
                        OperateResult<float[]> ret = cip.ReadFloat(ReadObject, length);
                        if (ret.IsSuccess)
                        {
                            for (int i = 0; i < input.Length; i++)
                            {
                                StationListlnfo.arrDataPoint[input[i].stationNumber - 1].strCellCode = tool.ConvertFloatArrayToAscii(ret.Content, 0 + 50 * i, 49 + 50 * i);

                                allDataReadfromCIP.BarCode[input[i].stationNumber - 1] = StationListlnfo.arrDataPoint[input[i].stationNumber - 1].strCellCode;   
                            }
                        }
                        else
                        {
                            logNet.WriteError("[CIP]", ReadObject + "读取失败");
                            //Console.WriteLine(ReadObject + "read failed");
                        }


                    }
                    break;

                case "Station_EarCode":
                    {
                        ushort length = 500; // 数组长度为硬编码，由Excel读出，不知后续需要是否需要更改
                        OperateResult<float[]> ret = cip.ReadFloat(ReadObject, length);
                        if (ret.IsSuccess)
                        {
                            for (int i = 0; i < input.Length; i++)
                            {
                                StationListlnfo.arrDataPoint[input[i].stationNumber - 1].strPoleEarCode = tool.ConvertFloatArrayToAscii(ret.Content, 0 + 50 * i, 49 + 50 * i);

                                allDataReadfromCIP.EarCode[input[i].stationNumber - 1] = StationListlnfo.arrDataPoint[input[i].stationNumber - 1].strPoleEarCode;
                            }
                        }
                        else
                        {
                            logNet.WriteError("[CIP]", ReadObject + "读取失败");
                            //Console.WriteLine(ReadObject + "read failed");
                        }
                    }
                    break;

                default:
                    break;

            }          
            
        }
   
        #endregion 
 
        #region Function 读取1000ms的数据 （功能开关，生产统计，报警信号，寿命管理, OEE)

        public bool[] ReadOneSecData(OneSecInfoStruct_CIP[] input, OmronConnectedCipNet cip)
        {           
            ushort length = (ushort)input.Length;
            var AlarmValue = new bool[length];
            if (input[0].varType == "BOOL" && input[0].varName != "Manual_Andon[10]")   //区分Manual_Andon 不连续数组
            {               
                OperateResult<bool[]> ret = cip.ReadBool(input[0].varName, length);

                if (ret.IsSuccess)
                {
                    AlarmValue = ret.Content;

                }
                else
                {
                    logNet.WriteError("[CIP]", input[0].varName + "读取失败");
                   // Console.WriteLine(input[0].varName + "read failed");
                }
            }
            else if (input[0].varName == "Manual_Andon[10]")
            {
                for (int i=0; i<input.Length;i++)
                {
                    OperateResult<bool> ret = cip.ReadBool(input[i].varName);
                    if (ret.IsSuccess)
                    {
                        AlarmValue[i] = ret.Content;
                    }
                    else
                    {
                        logNet.WriteError("[CIP]", input[0].varName + "读取失败");
                        //Console.WriteLine(input[i].varName + "read failed");
                    }
                    
                }
            }

            return AlarmValue;

        }
        public void ReadOneSecData(OneSecInfoStruct_CIP[] input, OmronConnectedCipNet cip, ref AllDataReadfromCIP allDataReadfromCIP, ref DeviceDataStruct_IEC DeviceDataStruct)
        {
       
            if (input[0].varType == "BOOL" )   //区分Manual_Andon 不连续数组
            {
                ushort length = (ushort)input.Length;
                OperateResult<bool[]> ret = cip.ReadBool(input[0].varName, length);
                if (ret.IsSuccess)
                {
                    Array.Copy(ret.Content, 0, DeviceDataStruct.Value_FE, 0, input.Length);
                    Array.Copy(ret.Content, 0, allDataReadfromCIP.FunctionEnableValue, 0, ret.Content.Length); //写入Excel
                }
                else
                {
                    logNet.WriteError("[CIP]", input[0].varName + "读取失败");
                    //Console.WriteLine(input[0].varName + "read failed");
                }
            }

            else if (input[0].varType == "DINT")
            {
                ushort length = (ushort)input.Length;
                OperateResult<int[]> ret = cip.ReadInt32(input[0].varName, length);

                if (ret.IsSuccess)
                {
                    switch (input[0].varName)
                    {
                        case "Production_statistics[0]":
                            {

                                Array.Copy(ret.Content, 0, allDataReadfromCIP.ProductionDataValue, 0, ret.Content.Length); //写入Excel
                                Array.Copy(ret.Content, 0, DeviceDataStruct.Value_PD, 0, ret.Content.Length); //写入 DeviceDataStruct 结构体

                            }

                            break;

                        case "Cutterused_statistics[0]":  //寿命管理需要转为UDINT
                            {
                                var tempArray = new uint[ret.Content.Length];
                                for (int i=0; i < ret.Content.Length;i++)
                                {
                                    tempArray[i] = (uint)ret.Content[i];
                                }

                                Array.Copy(tempArray, 0, allDataReadfromCIP.LifeManagementValue, 0,  tempArray.Length); //写入Excel

                                Array.Copy(tempArray, 0, DeviceDataStruct.Value_LM, 0, tempArray.Length); //写入 DeviceDataStruct 结构体
                            }
                            break;


                        default:
                            break;

                    }
                }
                else
                {
                    logNet.WriteError("[CIP]", input[0].varName + "读取失败");
                    //Console.WriteLine(input[0].varName + "read failed");
                }
            }         
        }
        
        #endregion



        //通过数组的起终索引，来发送子数组 （六大工位）
        public void WriteSubArray(StationInfoStruct_CIP[] input, ref AllDataReadfromCIP allDataReadfromCIP, float[] sourceArray, ref UDT_ProcessStationDataValue ProcessStationDataValue)
        {
            var senddata = new string[input.Length];
            var i = 0;  //写入数组中的索引

            //按照数组索引 把数据放入发送区
            for (int j = 0; j < input.Length; j++)
            {
                senddata[j] = sourceArray[input[j].varIndex].ToString();
            }

            // 根据所属工位号 判断数组的索引
            switch (input[0].stationName)
            {
                case "加工工位（冲膜）":                 
                    Array.Copy(senddata, allDataReadfromCIP.ChongMoValue, senddata.Length);  //写到数据暂存区
                    i = 0; 
                    break;

                case "加工工位（热压1）":
                    Array.Copy(senddata, allDataReadfromCIP.ReYaValue_1, senddata.Length);  //写到数据暂存区
                    i = 1;
                    break;

                case "加工工位（热压2）":
                    Array.Copy(senddata, allDataReadfromCIP.ReYaValue_2, senddata.Length);  //写到数据暂存区
                    i = 2;
                    break;

                case "加工工位（热压3）":
                    Array.Copy(senddata, allDataReadfromCIP.ReYaValue_3, senddata.Length);  //写到数据暂存区
                    i = 3;
                    break;

                case "加工工位（热压4）":
                    Array.Copy(senddata, allDataReadfromCIP.ReYaValue_4, senddata.Length);  //写到数据暂存区
                    i = 4;
                    break;

                case "加工工位（顶封）":
                    Array.Copy(senddata, allDataReadfromCIP.DingFengValue, senddata.Length);  //写到数据暂存区
                    i = 5;
                    break;

                case "加工工位（右角封）":
                    Array.Copy(senddata, allDataReadfromCIP.YouJiaoFengValue, senddata.Length);  //写到数据暂存区
                    i = 6;
                    break;

                case "加工工位（左角封）":
                    Array.Copy(senddata, allDataReadfromCIP.ZuoJiaoFengValue, senddata.Length);  //写到数据暂存区
                    i = 7;
                    break;

                case "加工工位（侧封）":
                    Array.Copy(senddata, allDataReadfromCIP.CeFengValue, senddata.Length);  //写到数据暂存区
                    i = 8;
                    break;

                default:
                    break;

            }


            ProcessStationDataValue.arrDataPoint[i].iDataCount = (short)senddata.Length;
          
            for (int j=0; j< senddata.Length; j++)
            {
                ProcessStationDataValue.arrDataPoint[i].arrDataPoint[j].StringValue = senddata[j];
            }
            //Array.Copy(senddata, 0, ProcessStationDataValue.arrDataPoint[i].arrDataPoint, 0, senddata.Length);


            //float[] ToSendArray = new float[input.Length]; //将var修改成了ToSendArray
            //var startIndex = input[0].varIndex;
            //var length = input.Length;

            //Array.Copy(sourceArray, startIndex, ToSendArray, 0, length);

            //string StationName_Now = CN2EN(input[0].stationName);
            //var listWriteItem = new List<WriteItem>();

            ////Console.WriteLine(" {0}[0] value = {1}", StationName_Now, ToSendArray[0].ToString());

            //CopyStationData(ref allDataReadfromCIP, ToSendArray, input[0].stationName);  //将数据写到暂存区




            //try
            //{
            //    listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary[StationName_Now], Arp.Type.Grpc.CoreType.CtArray, ToSendArray));
            //    //Console.WriteLine(ToSendArray.Length);
            //    var writeItemsArray = listWriteItem.ToArray();
            //    var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
            //    bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
            //    //Console.WriteLine("WriteDataToDataAccessService ({0}) : {1}", nodeidDictionary[StationName_Now], result);               
            //}
            //catch (Exception e)
            //{
            //    logNet.WriteError("[Grpc]", StationName_Now + " 数据发送失败：" + e);
            //    //Console.WriteLine("ERRO: {0}", e);
            //}         

        }

        //XML标签转换 工位结构体数组的工位名是中文，为了方便XML与字典对应，需要转化为英文
        private string CN2EN(string NameCN)
        {
            string NameEN = "";

            switch(NameCN)
            {
                case "加工工位（冲膜）":
                    NameEN = "chongmo";
                    break;

                case "加工工位（热压）":
                    NameEN = "reya";
                    break;

                case "加工工位（顶封）":
                    NameEN = "dingfeng";
                    break;

                case "加工工位（右角封）":
                    NameEN = "youjiaofeng";
                    break;

                case "加工工位（左角封）":
                    NameEN = "zuojiaofeng";
                    break;

                case "加工工位（侧封）":
                    NameEN = "cefeng";
                    break;

                default:
                    break;

            }

            return NameEN;
        }

        public void CopyStationData(ref AllDataReadfromCIP allDataReadfromCIP, float[] SourceData, string NameCN)
        {
            switch (NameCN)
            {
                case "加工工位（冲膜）":
                    Array.Copy(SourceData, allDataReadfromCIP.ChongMoValue, SourceData.Length);  //写到数据暂存区
                    break;

                case "加工工位（热压1）":
                    Array.Copy(SourceData, allDataReadfromCIP.ReYaValue_1, SourceData.Length);  //写到数据暂存区
                    break;

                case "加工工位（热压2）":
                    Array.Copy(SourceData, allDataReadfromCIP.ReYaValue_2, SourceData.Length);  //写到数据暂存区
                    break;

                case "加工工位（热压3）":
                    Array.Copy(SourceData, allDataReadfromCIP.ReYaValue_3, SourceData.Length);  //写到数据暂存区
                    break;

                case "加工工位（热压4）":
                    Array.Copy(SourceData, allDataReadfromCIP.ReYaValue_4, SourceData.Length);  //写到数据暂存区
                    break;

                case "加工工位（顶封）":
                    Array.Copy(SourceData, allDataReadfromCIP.DingFengValue, SourceData.Length);  //写到数据暂存区
                    break;

                case "加工工位（右角封）":
                    Array.Copy(SourceData, allDataReadfromCIP.YouJiaoFengValue, SourceData.Length);  //写到数据暂存区
                    break;

                case "加工工位（左角封）":
                    Array.Copy(SourceData, allDataReadfromCIP.ZuoJiaoFengValue, SourceData.Length);  //写到数据暂存区
                    break;

                case "加工工位（侧封）":
                    Array.Copy(SourceData, allDataReadfromCIP.CeFengValue, SourceData.Length);  //写到数据暂存区
                    break;

                default:
                    break;

            }



        }


        #region 读取点位名(三个函数重载)

        public void ReadPointName(OneSecInfoStruct_CIP[] InputStruct, ref OneSecPointNameStruct_IEC OneSecNameStruct)
        {
            switch (InputStruct[0].varName)
            {
                case "SYS_Manual[0]":             // 功能开关
                    {
                        OneSecNameStruct.DataCount_FE = InputStruct.Length;
                        for (int i = 0; i < InputStruct.Length; i++)
                        {
                            OneSecNameStruct.Name_FE[i].StringValue = InputStruct[i].varAnnotation;
                        }
                    }
                    break;

                case "Production_statistics[0]":  // 生产统计
                    {
                        OneSecNameStruct.DataCount_PD = InputStruct.Length;
                        for (int i = 0; i < InputStruct.Length; i++)
                        {
                            OneSecNameStruct.Name_PD[i].StringValue = InputStruct[i].varAnnotation;
                        }
                    }
                    break;

                case "Cutterused_statistics[0]":  // 寿命管理
                    {
                        OneSecNameStruct.DataCount_LM = InputStruct.Length;
                        for (int i = 0; i < InputStruct.Length; i++)
                        {
                            OneSecNameStruct.Name_LM[i].StringValue = InputStruct[i].varAnnotation;
                        }
                    }
                    break;

                default:
                    break;
            }

           

        }
        public void ReadPointName(OneSecInfoStruct_CIP[] Y6, OneSecInfoStruct_CIP[] Manual_Andon, ref OneSecPointNameStruct_IEC OneSecNameStruct)
        {
            OneSecNameStruct.DataCount_OEE = Y6.Length + Manual_Andon.Length;
        
            for (int i = 0; i < Y6.Length; i++)
            {
                OneSecNameStruct.Name_OEE[i].StringValue = Y6[i].varAnnotation;
            }
            for (int i = 0; i < Manual_Andon.Length; i++)
            {              
                OneSecNameStruct.Name_OEE[i + Y6.Length].StringValue = Manual_Andon[i].varAnnotation;
            }
        }      
        public void ReadPointName(List<OneSecInfoStruct_CIP[]> alarmGroups, int stringnumber, ref OneSecPointNameStruct_IEC OneSecNameStruct)
        {
            var index = 0;
            OneSecNameStruct.DataCount_ALM = stringnumber;
            foreach (var alarmGroup in alarmGroups)
            {
                foreach (var alarm in alarmGroup)
                {
                    OneSecNameStruct.Name_ALM[index++].StringValue = alarm.varAnnotation;
                }
            }
        }


        public void ReadPointName(List<StationInfoStruct_CIP[]> StationDataStruct, ref ProcessStationNameStruct_IEC ProcessStationNameStruct)
        {
            ProcessStationNameStruct.StationCount = (short)StationDataStruct.Count;   //写入加工工位的个数

            var i = 0;  //工位数量的索引
            
            foreach (var StationData in StationDataStruct)
            {
                ProcessStationNameStruct.UnitStation[i].DataCount = (short)StationData.Length;   //每个加工工位的点位数量 （不超过16个点位）
                ProcessStationNameStruct.UnitStation[i].StationNO = (short)StationData[0].StationNumber;
                ProcessStationNameStruct.UnitStation[i].StationName = StationData[0].stationName;

                var j = 0;  //每个工位里采集值的索引
                foreach (var item in StationData)
                {
                    ProcessStationNameStruct.UnitStation[i].arrKey[j].StringValue = item.varAnnotation;
                    j++;
                }
                i++;
            }
        }

        #endregion

        #region 

        public void ReadandSendStaionInfo(DeviceInfoConSturct_CIP[] InputStruct, GrpcTool grpcToolInstance, Dictionary<string, string> nodeidDictionary, IDataAccessServiceClient grpcDataAccessServiceClient, CallOptions options1)
        {
            var senddata = new StationInfoStruct_IEC();

            senddata.StationCount = (short)InputStruct.Length; 

            for(int i = 0;i<InputStruct.Length;i++)
            {
                senddata.NextStationNO[i] = (short)InputStruct[i].nextStationNumber;
                senddata.TempCodeNO[i] = (short)InputStruct[i].pseudoCode;
            }

            var listWriteItem = new List<WriteItem>();

            try
            {
                listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary.GetValueOrDefault("StationInfoStruct"), Arp.Type.Grpc.CoreType.CtStruct, senddata));
                var writeItemsArray = listWriteItem.ToArray();
                var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
            }
            catch (Exception e)
            {
                logNet.WriteError("[Grpc]", " 设备信息表中的信息发送失败：" + e);
                //Console.WriteLine("ERRO: {0}", e);
            }


        }

        #endregion






    }
}

