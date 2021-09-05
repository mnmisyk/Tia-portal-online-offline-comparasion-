#define V16
#define Plcsim
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Win32;
using Siemens.Engineering;
using Siemens.Engineering.Compiler;
using Siemens.Engineering.Connection;
using Siemens.Engineering.Hmi;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.Online;
using Siemens.Engineering.SW;
using NetworkInterface = System.Net.NetworkInformation.NetworkInterface;
using System.Configuration;
using System.Reflection;

namespace ShitVersionDog
{
    class Program
    {
        static List<FileSystemInfo> FsInfoList = new List<FileSystemInfo>();
        static string NetworkCard;
        static int _MAXPOOL = 0;
        static Queue<string> FileQueue = new Queue<string>();
        //源文件夹
        static string SourceFolder = ConfigurationManager.AppSettings["SourceFolder"];

        //目标生成文件夹
        static string TargetFolder = ConfigurationManager.AppSettings["TargetFolder"];
        static string LogPlcPath = TargetFolder + @"\log_PLC.txt";


        static void Main(string[] args)
        {
            //test tia openness availability
            //TiaPortal _tiaPortal = new TiaPortal(TiaPortalMode.WithoutUserInterface);
         

            DelectDir(TargetFolder);
            List<string> fl = new List<string>();   //最后实际出来的zap14 ，最新的项目
            List<string> Zap16Files = new List<string>();

            #region 交互
            Console.Clear();
            Console.WriteLine("###################################################");
            Console.WriteLine("STEP 1 ,choose the interface which can conect plc ");
            Console.WriteLine("###################################################");
            NetworkInterface[] adapters = NetworkInterface.GetAllNetworkInterfaces();//获取本地计算机上网络接口的对象
            //Console.WriteLine("network cards count：" + adapters.Length);
            Console.WriteLine();
            int j = 1;
            List<NetworkInterface> ActiveInterfaceList = new List<NetworkInterface>();
            foreach (NetworkInterface adapter in adapters)
            {
                if (adapter.OperationalStatus == OperationalStatus.Up)
                {
                    ActiveInterfaceList.Add(adapter);
                    Console.WriteLine((j++) + " ,Description：" + adapter.Description);
                    IPInterfaceProperties property = adapter.GetIPProperties();
                    foreach (UnicastIPAddressInformation ip in property.UnicastAddresses)
                    {
                        if (ip.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                        {
                            Console.WriteLine(ip.Address.ToString());
                        }
                    }

                    // 格式化显示MAC地址                
                    PhysicalAddress pa = adapter.GetPhysicalAddress();//获取适配器的媒体访问（MAC）地址
                    byte[] bytes = pa.GetAddressBytes();//返回当前实例的地址
                    StringBuilder sb = new StringBuilder();
                    for (int i = 0; i < bytes.Length; i++)
                    {
                        sb.Append(bytes[i].ToString("X2"));//以十六进制格式化
                        if (i != bytes.Length - 1)
                        {
                            sb.Append("-");
                        }
                    }
                    Console.WriteLine("MAC address：" + sb);
                }

                //Console.WriteLine("标识符：" + adapter.Id);
                //Console.WriteLine("名称：" + adapter.Name);
                //Console.WriteLine("类型：" + adapter.NetworkInterfaceType);
                //Console.WriteLine("速度：" + adapter.Speed * 0.001 * 0.001 + "M");
                //Console.WriteLine("操作状态：" + adapter.OperationalStatus);
                //Console.WriteLine("MAC 地址：" + adapter.GetPhysicalAddress());


                Console.WriteLine();
            }
            Console.WriteLine("choose from 1 to " + (--j) + " :");
            string x = Console.ReadLine();
            NetworkCard = ActiveInterfaceList[Convert.ToInt32(x) - 1].Description;
            Console.WriteLine("< " + NetworkCard + " > is selected ;");



            Console.WriteLine();
            Console.WriteLine("#####################################################################");
            Console.WriteLine("Step 2");
            Console.WriteLine("#####################################################################");
            Console.WriteLine("Choose how many TIA instances you wanna run at same time ,which depends on your PC performance");
            Console.WriteLine("16GB RAM ,I7 8 cores ,suggestion value would be 5  ,default Value 3");
            int MaxThreads = Convert.ToInt16(Console.ReadLine());
           // Console.WriteLine(@"fill project directory ,type enter to use " + @" D:\TMP_Files\Projects");
           // string ProjectPath = Console.ReadLine();
           // if (ProjectPath.Trim().Length == 0)
           // {
             string   ProjectPath = TargetFolder;
            // }
            // Console.ReadKey();

            #endregion
            director(SourceFolder, ref fl, ".zap16");
            //获取到实际的文件信息后再利用linq筛选
            var query = from f in FsInfoList
                        group f by f.Name.Split('_')[0] into g
                        select g.OrderByDescending(p => p.LastAccessTime); //按修改时间排序,寻找最新的项目
            foreach (var item in query)
            {
                var k = item.First();
                if (!Zap16Files.Contains(k.FullName))
                {
                    Zap16Files.Add(k.FullName);
                }
            }


            _7zHelper.RunCMDCommand(Zap16Files, TargetFolder).Wait();

            #region get ap14 path
            List<string> FileList = new List<string>();
            List<string> FileListAp16 = new List<string>();
            director(ProjectPath, ref FileList, ".ap16");
            foreach (var item in FileList)
            {
                if (item.Substring(item.Length - 5) == ".ap16")
                {
                    FileListAp16.Add(item);
                }
            }
            #endregion
            File.Delete(LogPlcPath);

            foreach (var item in FileListAp16)
            {
                FileQueue.Enqueue(item);
            }


            Task.Run(() =>
            {
                while (true)
                {
                    if (_MAXPOOL != 0 || FileQueue.Count != 0)
                    {
                        Console.WriteLine("thread count: " + _MAXPOOL + " ,and " + FileQueue.Count + " projects left");
                    }
                    
                    if (_MAXPOOL < MaxThreads)
                    {
                        if (FileQueue.Count > 0)
                        {
                            Thread.Sleep(2000);
                            string FilePath = FileQueue.Dequeue();
                            Task<string> OnlineCompare = onlineAsync(FilePath);
                        }
                    }

                    if (_MAXPOOL==0 && FileQueue.Count==0)
                    {
                        Console.WriteLine("finish compare ");

                    }

                }

            });



            Console.ReadKey();



        }
        static ReaderWriterLockSlim LogWriteLock = new ReaderWriterLockSlim();



        public static List<string> director(string dirs, ref List<string> fs, string extension)
        {
            //定义一个list集合
            // List<FileSystemInfo> FileList = new List<FileSystemInfo>();

            //绑定到指定的文件夹目录
            DirectoryInfo dir = new DirectoryInfo(dirs);
            //检索表示当前目录的文件和子目录
            FileSystemInfo[] fsinfos = dir.GetFileSystemInfos();
            //遍历检索的文件和子目录
            foreach (FileSystemInfo fsinfo in fsinfos)
            {
                //判断是否为空文件夹　　
                if (fsinfo is DirectoryInfo)
                {
                    //递归调用
                    director(fsinfo.FullName, ref fs, extension);
                }
                else
                {
                    var stationName = fsinfo.Name.Split('_')[0];
                    if (fsinfo.Extension == extension)
                    {
                        FsInfoList.Add(fsinfo);
                        Console.WriteLine(fsinfo.FullName);
                        fs.Add(fsinfo.FullName);

                    }

                }
            }




            return fs;
        }

        public static async Task<string> onlineAsync(string path)
        {
            Interlocked.Increment(ref _MAXPOOL);
            string x = await CompareOnlieResult(path);
            Interlocked.Decrement(ref _MAXPOOL);
            return x;
        }

        private static  Task<string> CompareOnlieResult(string file_path)
        {
            

            return Task.Run(() =>
            {

               
                string Result = "Unable to compare";
                //Console.WriteLine(Result);
                TiaPortal _tiaPortal = new TiaPortal(TiaPortalMode.WithoutUserInterface);
                try
                {

                    Project _tiaPortalProject = _tiaPortal.Projects.Open(new FileInfo(file_path));
                    foreach (var device in _tiaPortalProject.Devices)
                    {
                        var DeviceIs_PlcSoftware = GetPlcSoftware(device);
                        //WriteStatusEntry(device.Name + ": is PlcSoftware ? " + (DeviceIs_PlcSoftware is PlcSoftware));
                        if (DeviceIs_PlcSoftware is PlcSoftware)
                        {
                            var deviceitem = DeviceIs_PlcSoftware.Parent.Parent as DeviceItem;
                            var onlineTarget = deviceitem.GetService<OnlineProvider>();
                            var configuration = onlineTarget.Configuration;
                            Console.WriteLine("TypeIdentifier of PlcSoftware :" + deviceitem.TypeIdentifier);
                            SetConnectionWithSlot(onlineTarget); //Intel(R) Ethernet Connection (2) I219-LM
                            if (onlineTarget.Configuration.IsConfigured)
                            {
                                Console.WriteLine("online status:" + onlineTarget.State);
                                if (onlineTarget.State == OnlineState.Offline)
                                {
                                    onlineTarget.GoOnline();
                                }

                            }
                            else
                            {
                                Result = ("Unable to compare ,NEED online Configuratopn");
                            }


                            if (onlineTarget.State == OnlineState.Online)
                            {

                                // WriteStatusEntry(@"Comparing PLC: \n");
                                var compareResult = (DeviceIs_PlcSoftware as PlcSoftware).CompareToOnline();

                                Result = "  Result:" + compareResult.RootElement.ComparisonResult + "   " + "DetailedInfo:" + compareResult.RootElement.DetailedInformation;
                                onlineTarget.GoOffline();
                            }
                            else
                            {
                                Result = "Unable to compare, PLC is not connected. Please connect PLC first.";
                            }
                            LogWriteLock.EnterWriteLock();
                            File.AppendAllText(LogPlcPath, "\r\n\r\n" + DateTime.Now.ToString() + "   " + file_path.Split('\\')[file_path.Split('\\').Length - 1] + "  " + Result);
                            LogWriteLock.ExitWriteLock();

                        }

                    }

                }
                catch (Exception exc)
                {
                    LogWriteLock.EnterWriteLock();
                    File.AppendAllText(LogPlcPath, "\r\n\r\n" + DateTime.Now.ToString() + "   " + file_path.Split('\\')[file_path.Split('\\').Length - 1] + "  " + Result + "   " + exc.Message);
                    LogWriteLock.ExitWriteLock();
                }
                finally
                {
                    _tiaPortal.Dispose();

                }

                
                return Result;
            });
           
        }


        public static PlcSoftware GetPlcSoftware(Device device)
        {
            if (device == null)
                throw new ArgumentNullException(nameof(device), "Parameter is null");
            //if (device.Subtype.ToLower().Contains("sinamics"))
            //    return null;

            var itemStack = new Stack<DeviceItem>();
            foreach (var item in device.DeviceItems)
            {
                itemStack.Push(item);
            }

            while (itemStack.Count != 0)
            {
                var item = itemStack.Pop();

                var target = item.GetService<SoftwareContainer>();
                if (target != null && target.Software is PlcSoftware)
                {
                    return (PlcSoftware)target.Software;
                }

                foreach (var subItem in item.DeviceItems)
                {
                    itemStack.Push(subItem);
                }
            }

            return null;
        }

        public static void SetConnectionWithSlot(OnlineProvider onlineProvider)
        {
            ConnectionConfiguration configuration = onlineProvider.Configuration;
            ConfigurationMode mode = configuration.Modes.Find(@"PN/IE");
#if Plcsim
            ConfigurationPcInterface pcInterface = mode.PcInterfaces.Find("PLCSIM", 1);
#else
 ConfigurationPcInterface pcInterface = mode.PcInterfaces.Find(NetworkCard, 1);
#endif

            // or network pc interface that is connected to plc

            ConfigurationTargetInterface slot = pcInterface.TargetInterfaces.Find("1 X1");
            configuration.ApplyConfiguration(slot);
            // After applying configuration, you can go online
            onlineProvider.GoOnline();
        }


        public static void DelectDir(string srcPath)
        {
            if (Directory.Exists(srcPath))
            {
                try
                {
                    DirectoryInfo dir = new DirectoryInfo(srcPath);
                    FileSystemInfo[] fileinfo = dir.GetFileSystemInfos();  //返回目录中所有文件和子目录
                    foreach (FileSystemInfo i in fileinfo)
                    {
                        if (i is DirectoryInfo)            //判断是否文件夹
                        {
                            DirectoryInfo subdir = new DirectoryInfo(i.FullName);
                            subdir.Delete(true);          //删除子目录和文件
                        }
                        else
                        {
                            File.Delete(i.FullName);      //删除指定文件
                        }
                    }
                }
                catch (Exception e)
                {
                    throw;
                }
            }
        }

    }
}
