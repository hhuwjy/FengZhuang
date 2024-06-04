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




namespace Ph_CipComm_FengZhuang
{

    class OmronComm
    {

        //#region Function 读取六个工位的数据
        //public void ReadandSendStation(StationInfoStruct_CIP[] input, OmronConnectedCipNet cip , GrpcTool grpcToolInstance, Dictionary<string, string> nodeidDictionary, IDataAccessServiceClient grpcDataAccessServiceClient, CallOptions options1)
        //{
        //    //var tempstring = "";  //暂存取到的string数据
        //    //int count = 0; //计数器
        //    ushort length = (ushort)input.Length;
        //    string StationName_Now = CN2EN(input[0].stationName); //将当前结构体数组的工位名读取出来，后续在xml文件中对应,中文转拼音（英文）
        //    var listWriteItem = new List<WriteItem>();
        //    //WriteItem[] writeItems = new WriteItem[] { };

        //    listWriteItem.Clear();  //每次发送工位数据前都清空list

        //    OperateResult<float[]> ret = cip.ReadFloat(input[0].varName, length);
        //    if (ret.IsSuccess)
        //    {
        //        //writeItems = null;
        //        try
        //        {
        //            listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary.GetValueOrDefault(StationName_Now), Arp.Type.Grpc.CoreType.CtArray, ret.Content)); //todo:待优化floatArr改为Content
        //            var writeItemsArray = listWriteItem.ToArray();
        //            var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
        //            bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
        //        }
        //        catch (Exception e)
        //        {
        //            Console.WriteLine("ERRO: {0}", e);
        //        }

        //        //SendDataToIEC(listWriteItem);
                
        //    }
        //    else
        //    {
        //        //logNet.WriteInfo(input[0].varName + "read failed");
        //        Console.WriteLine(input[0].varName + "read array failed");

        //    }

        //}
        //#endregion  Function 读取六个工位的数据


        #region Function 读取设备信息（以数组形式一起读上来，再按照序号写入对应的工位里）
        public void ReadDeviceInfoConSturct1(DeviceInfoConSturct1_CIP[] input, OmronConnectedCipNet cip, StringBuilder[] Output)
        {
            string ReadObject = input[0].varName;   //！这里约定变量名就叫Auto_process 索引单独是一个变量！
            ushort Auto_Process_Length = 86;  // 数组长度为硬编码，由Excel读出，不知后续需要是否需要更改
            ushort Clear_Manual = 76;
            ushort Cell = 57;
            ushort Code = 500;
           
            if (input[0].varType == "DINT")
            {
                OperateResult<int[]> ret = cip.ReadInt32(ReadObject, Auto_Process_Length);
                if (ret.IsSuccess)
                {
                    for (int i = 0; i < input.Length; i++)
                    {               
                        Output[input[i].stationNumber].Append(ret.Content[input[i].varIndex].ToString() + ",");
                    }
                }
                else
                {
                    //logNet.WriteInfo(ReadObject + "read failed");
                    Console.WriteLine(ReadObject + "read failed");

                }
            }

            else if (input[0].varType == "BOOL")
            {
                ushort length = input[0].varName == "Clear_Manual" ? Clear_Manual : Cell;
                OperateResult<bool[]> ret = cip.ReadBool(ReadObject, length);
                if (ret.IsSuccess)
                {
                    for (int i = 0; i < input.Length; i++)
                    {
                        Output[input[i].stationNumber].Append(ret.Content[input[i].varIndex] ? "1," : "0,");
                    }
                }
                else
                {
                    logNet.WriteInfo(ReadObject + "read failed");
                    //Console.WriteLine(ReadObject + "read failed");
                }
                  
            }

            else if (input[0].varType == "REAL")
            {
                OperateResult<float[]> ret = cip.ReadFloat(ReadObject, Code);
                if (ret.IsSuccess)
                {
                    for (int i = 0; i < input.Length; i++)
                    {
                        var tempstring = string.Join("", Output, 0 + 50 * i, 49 + 50 * i);
                        float[] output_temp = new float[50];
                        Output[input[i].stationNumber] = tool.ConvertFloatArrayToAscii(ret.Content, 0 + 50 * i, 49 + 50 * i);                        
                    }
                }
                else
                {
                    logNet.WriteInfo(ReadObject + "read failed");
                    //Console.WriteLine(ReadObject + "read failed");

                }
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
                    logNet.WriteInfo(input[0].varName + "read failed");
                   // Console.WriteLine(input[0].varName + "read failed");
                }
            }
            else if (input[0].varName == "Manual_Andon[10]")
            {
                for (int i=0; i<input.Length;i++)
                {
                    OperateResult<bool> ret = cip.ReadBool(input[0].varName);
                    if (ret.IsSuccess)
                    {
                        AlarmValue[i] = ret.Content;
                    }
                    else
                    {
                        logNet.WriteInfo(input[0].varName + "read failed");
                        //Console.WriteLine(input[i].varName + "read failed");
                    }
                    
                }
            }

            return AlarmValue;

        }
        public void ReadandSendOneSecData(OneSecInfoStruct_CIP[] input, OmronConnectedCipNet cip, int IECNumber,GrpcTool grpcToolInstance, Dictionary<string, string> nodeidDictionary, IDataAccessServiceClient grpcDataAccessServiceClient, CallOptions options1)
        {
            var listWriteItem = new List<WriteItem>();
            //WriteItem[] writeItems = new WriteItem[] { };

            if (input[0].varType == "BOOL" )   //区分Manual_Andon 不连续数组
            {
                ushort length = (ushort)input.Length;
                OperateResult<bool[]> ret = cip.ReadBool(input[0].varName, length);
                var senddata =new bool[IECNumber];
                if (ret.IsSuccess)
                {                  
                    if (input.Length< IECNumber)
                    {
                        Array.Copy(ret.Content, 0, senddata, 0, input.Length);
                        try
                        {
                           listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary.GetValueOrDefault(input[0].varName), Arp.Type.Grpc.CoreType.CtArray, senddata));
                            var writeItemsArray = listWriteItem.ToArray();
                            var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                            bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("ERRO: {0}", e);
                        }
                    }
                    else
                    {
                        try
                        {
                            listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary.GetValueOrDefault(input[0].varName), Arp.Type.Grpc.CoreType.CtArray, ret.Content));
                            var writeItemsArray = listWriteItem.ToArray();
                            var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                            bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("ERRO: {0}", e);
                        }
                    }

                }
                else
                {
                    //logNet.WriteInfo(input[0].varName + "read failed");
                    Console.WriteLine(input[0].varName + "read failed");
                }
            }

            else if (input[0].varType == "DINT")
            {
                ushort length = (ushort)input.Length;
                OperateResult<int[]> ret = cip.ReadInt32(input[0].varName, length);
                var senddata = new int[IECNumber];
                if (ret.IsSuccess)
                {
                    if (input.Length < IECNumber)
                    {
                        Array.Copy(ret.Content, 0, senddata, 0, input.Length);
                        try
                        {
                            listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary.GetValueOrDefault(input[0].varName), Arp.Type.Grpc.CoreType.CtArray, senddata));
                            var writeItemsArray = listWriteItem.ToArray();
                            var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                            bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("ERRO: {0}", e);
                        }

                    }
                    else
                    {
                        try
                        {
                            listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary.GetValueOrDefault(input[0].varName), Arp.Type.Grpc.CoreType.CtArray, ret.Content));
                            var writeItemsArray = listWriteItem.ToArray();
                            var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                            bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("ERRO: {0}", e);
                        }

                    }
                }
                else
                {
                    logNet.WriteInfo(input[0].varName + "read failed");
                    //Console.WriteLine(input[0].varName + "read failed");
                }
            }

            listWriteItem.Clear();
        }
        #endregion



        //通过数组的起终索引，来发送子数组
        public void SendSubArray(StationInfoStruct_CIP[] input, float[] sourceArray, GrpcTool grpcToolInstance, Dictionary<string, string> nodeidDictionary, IDataAccessServiceClient grpcDataAccessServiceClient , CallOptions options1)
        {
            float[] ToSendArray = new float[input.Length]; //将var修改成了ToSendArray
            var startIndex = input[0].varIndex;
            var length = input.Length;

            Array.Copy(sourceArray, startIndex, ToSendArray, 0, length);

            string StationName_Now = CN2EN(input[0].stationName);
            var listWriteItem = new List<WriteItem>();

            //Console.WriteLine(" {0}[0] value = {1}", StationName_Now, ToSendArray[0].ToString());

            try
            {
                listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary[StationName_Now], Arp.Type.Grpc.CoreType.CtArray, ToSendArray));
                //Console.WriteLine(ToSendArray.Length);
                var writeItemsArray = listWriteItem.ToArray();
                var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
                //Console.WriteLine("WriteDataToDataAccessService ({0}) : {1}", nodeidDictionary[StationName_Now], result);
                 
            }

            catch (Exception e)
            {
                Console.WriteLine("ERRO: {0}", e);
            }         
          
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

        //读取和发送点位名(三个函数重载)
        public void ReadandSendPointName(OneSecInfoStruct_CIP[] InputStruct, OneSecPointNameStruct_IEC functionEnableNameStruct_IEC, int IEC_Array_Number, GrpcTool grpcToolInstance, Dictionary<string, string> nodeidDictionary, IDataAccessServiceClient grpcDataAccessServiceClient, CallOptions options1)
        {
            var listWriteItem = new List<WriteItem>();
            WriteItem[] writeItems = new WriteItem[] { };
            functionEnableNameStruct_IEC.iDataCount = InputStruct.Length;
            functionEnableNameStruct_IEC.stringArrData = new stringStruct[IEC_Array_Number];
            for (int i = 0; i < IEC_Array_Number; i++)
            {
                if (i < InputStruct.Length)
                {
                    functionEnableNameStruct_IEC.stringArrData[i].str = InputStruct[i].varAnnotation;
                }
                else
                {
                    functionEnableNameStruct_IEC.stringArrData[i].str = " ";
                }
            }
            try
            {
                listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary.GetValueOrDefault(InputStruct[0].varAnnotation), Arp.Type.Grpc.CoreType.CtStruct, functionEnableNameStruct_IEC));
                var writeItemsArray = listWriteItem.ToArray();
                var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);

            }
            catch (Exception e)
            {
                logNet.WriteError(nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff") + InputStruct[0].varAnnotation + "ERRO: {0}", e.ToString());
                //Console.WriteLine("ERRO: {0}", e);
            }

        }
        public void ReadandSendPointName(String[] InputString, OneSecPointNameStruct_IEC functionEnableNameStruct_IEC, int IEC_Array_Number, GrpcTool grpcToolInstance, Dictionary<string, string> nodeidDictionary, IDataAccessServiceClient grpcDataAccessServiceClient, CallOptions options1)
        {
            var listWriteItem = new List<WriteItem>();
            WriteItem[] writeItems = new WriteItem[] { };
            functionEnableNameStruct_IEC.iDataCount = InputString.Length;
            functionEnableNameStruct_IEC.stringArrData = new stringStruct[IEC_Array_Number];
            for (int i = 0; i < IEC_Array_Number; i++)
            {
                if (i < InputString.Length)
                {
                    functionEnableNameStruct_IEC.stringArrData[i].str = InputString[i];
                }
                else
                {
                    functionEnableNameStruct_IEC.stringArrData[i].str = " ";
                }
            }
            try
            {
                listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary.GetValueOrDefault(InputString[0]), Arp.Type.Grpc.CoreType.CtStruct, functionEnableNameStruct_IEC));
                var writeItemsArray = listWriteItem.ToArray();
                var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);

            }
            catch (Exception e)
            {
                logNet.WriteError(nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff") + InputString[0] + "ERRO: {0}", e.ToString());
                //Console.WriteLine("ERRO: {0}", e);
            }

        }
        public void ReadandSendPointName(StationInfoStruct_CIP[] InputStruct, OneSecPointNameStruct_IEC functionEnableNameStruct_IEC, int IEC_Array_Number, GrpcTool grpcToolInstance, Dictionary<string, string> nodeidDictionary, IDataAccessServiceClient grpcDataAccessServiceClient, CallOptions options1)
        {
            var listWriteItem = new List<WriteItem>();
            WriteItem[] writeItems = new WriteItem[] { };
            functionEnableNameStruct_IEC.iDataCount = InputStruct.Length;
            functionEnableNameStruct_IEC.stringArrData = new stringStruct[IEC_Array_Number];
            for (int i = 0; i < IEC_Array_Number; i++)
            {
                if (i < InputStruct.Length)
                {
                    functionEnableNameStruct_IEC.stringArrData[i].str = InputStruct[i].varAnnotation;
                }
                else
                {
                    functionEnableNameStruct_IEC.stringArrData[i].str = " ";
                }
            }
            try
            {            
                listWriteItem.Add(grpcToolInstance.CreatWriteItem(nodeidDictionary.GetValueOrDefault(InputStruct[0].varAnnotation), Arp.Type.Grpc.CoreType.CtStruct, functionEnableNameStruct_IEC));
                var writeItemsArray = listWriteItem.ToArray();
                var dataAccessServiceWriteRequest = grpcToolInstance.ServiceWriteRequestAddDatas(writeItemsArray);
                bool result = grpcToolInstance.WriteDataToDataAccessService(grpcDataAccessServiceClient, dataAccessServiceWriteRequest, new IDataAccessServiceWriteResponse(), options1);
            }
            catch (Exception e)
            {
                logNet.WriteError(nowDisplay.ToString("yyyy-MM-dd HH:mm:ss:fff") + InputStruct[0].varAnnotation + "ERRO: {0}", e.ToString());
                //Console.WriteLine("ERRO: {0}", e);
            }
          
        }




    }
}

