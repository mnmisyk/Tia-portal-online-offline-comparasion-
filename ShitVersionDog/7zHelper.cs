using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ShitVersionDog
{
    public static class _7zHelper
    {

        public static async Task RunCMDCommand(List<string> Zap14List, string targefolder)
        {
            int num = 0;
            var ComList = GetCommandList(Zap14List, targefolder);
            await Task.Run(() =>
            {
                foreach (string com in ComList)
                {
                   Task.Run(() =>
                   {
                       using (Process pc = new Process())
                       {
                           pc.StartInfo.FileName = "cmd.exe";
                           pc.StartInfo.CreateNoWindow = true;//隐藏窗口运行
                           pc.StartInfo.RedirectStandardError = true;//重定向错误流
                           pc.StartInfo.RedirectStandardInput = true;//重定向输入流
                           pc.StartInfo.RedirectStandardOutput = true;//重定向输出流
                           pc.StartInfo.UseShellExecute = false;
                           pc.Start();
                           // int lenght = command.Length;

                           pc.StandardInput.WriteLine(com);//输入CMD命令

                           pc.StandardInput.WriteLine("exit");//结束执行，很重要的
                           pc.StandardInput.AutoFlush = true;

                           var x = pc.StandardOutput.ReadToEnd();//读取结果 

                           pc.WaitForExit();
                           pc.Close();

                           Interlocked.Increment(ref num);
                           Console.WriteLine(x);
                           //Console.Clear();
                       }

                   });
                }


            });
            //to ensure all files are unpacked already .
            await Task.Run(() =>
            {
                while (num != Zap14List.Count)
                {
                    //do nothing ,just halt here
                }
                if (num == Zap14List.Count)
                {
                    num = 0;
                }

            });



        }
        private static List<string> GetCommandList(List<string> Zap14List, string TargetFolder)
        {


            List<string> CommandList = new List<string>();
            foreach (var Zap14 in Zap14List)
            {
                string com = @"c: && cd C:\Program Files\7-Zip && 7z.exe x """ + Zap14 + @"""  -o""" + TargetFolder + @"\" + Zap14.Split('\\')[Zap14.Split('\\').Length - 1].Split('.')[0] + @""" -aoa";
                CommandList.Add(com);
            }

            return CommandList;
        }

    }
}
