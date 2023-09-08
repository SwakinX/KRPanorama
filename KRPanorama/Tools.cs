using Sunny.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace KRPanorama
{
    internal class Tools
    {
        List<FileSystemInfo> fes = new List<FileSystemInfo>();
        


        /// <summary>
        /// 遍历查找html和xml文件
        /// </summary>
        /// <param name="dir"></param>
        /// <returns></returns>
        public List<FileSystemInfo> fileDirector(string dir)//遍历器
        {
            DirectoryInfo d = new DirectoryInfo(dir);
            FileSystemInfo[] fsinfos = d.GetFileSystemInfos();
            
            foreach (FileSystemInfo fsinfo in fsinfos)
            {
                if (fsinfo is DirectoryInfo)     //判断是否为文件夹
                {
                    fileDirector(fsinfo.FullName);//递归调用
                }
                else
                {
                    if (fsinfo.Name == "tour.html" || fsinfo.Name == "tour.xml")
                    {
                        fes.Add(fsinfo);
                    } 
                }
            }
            return fes;
        }
        /// <summary>
        /// 执行切片命令
        /// </summary>
        /// <param name="file"></param>
        public void bat(string path)
        {
            
            Process process = new Process();
            process.StartInfo.FileName = "cmd.exe"; //设置需要执行的命令
            process.StartInfo.Arguments = "/c " + path; //设置命令参数
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true; //重定向标准输出
            process.StartInfo.CreateNoWindow = true; //不显示窗口
            process.Start();
            string output = process.StandardOutput.ReadToEnd(); //读取标准输出
            process.WaitForExit(); //等待执行完毕
            
        }
        /// <summary>
        /// 修改全景信息
        /// </summary>
        /// <param name="file"></param>
        public void reple(FileInfo file)
        {
            
            string name = file.Directory.Parent.Name;//获取名称
            //BinaryReader binary = new BinaryReader((Stream)file);
            if (file.Extension is ".html")
            {
                HtmlAgilityPack.HtmlDocument document = new HtmlAgilityPack.HtmlDocument();
                document.Load(file.FullName);
                var node = document.DocumentNode.SelectNodes("//title");
                foreach (var nod in node)
                {
                    var title = nod.GetAttributeValue("title", string.Empty);
                    if (string.IsNullOrWhiteSpace(title))
                    {
                        nod.InnerHtml = name;
                    }
                }
                file.CopyTo(file.FullName + "-" + DateTime.Now.ToString("yyyy.MM.dd-HH.mm.ss") + ".html",true);
                document.Save(file.FullName);
            }
            if (file.Extension is ".xml")
            {
                XmlDocument xml = new XmlDocument();
                xml.Load(file.FullName);
                XmlElement el = xml.DocumentElement;
                el.SetAttribute("title", name);
                int i = 1;
                foreach (var item in el)
                {
                    if (item is XmlElement)
                    {
                        XmlElement e = item as XmlElement;
                        if (e.Name is "scene")
                        {
                            e.SetAttribute("title", "全景"+i);
                            i++;
                        }
                    }
                }
                file.CopyTo(file.FullName + "-" + DateTime.Now.ToString("yyyy.MM.dd-HH.mm.ss") + ".xml",true);
                xml.Save(file.FullName);
            }
            
        }
        /// <summary>
        /// 复制文件夹
        /// </summary>
        /// <param name="sourceDir"></param>
        /// <param name="destinationDir"></param>
        /// <param name="recursive"></param>
        /// <exception cref="DirectoryNotFoundException"></exception>
        static void CopyDirectory(string sourceDir, string destinationDir, bool recursive)
        {
            // Get information about the source directory
            var dir = new DirectoryInfo(sourceDir);

            // Check if the source directory exists
            if (!dir.Exists)
                throw new DirectoryNotFoundException($"Source directory not found: {dir.FullName}");

            // Cache directories before we start copying
            DirectoryInfo[] dirs = dir.GetDirectories();

            // Create the destination directory
            Directory.CreateDirectory(destinationDir);

            // Get the files in the source directory and copy to the destination directory
            foreach (FileInfo file in dir.GetFiles())
            {
                string targetFilePath = Path.Combine(destinationDir, file.Name);
                file.CopyTo(targetFilePath);
            }

            // If recursive and copying subdirectories, recursively call this method
            if (recursive)
            {
                foreach (DirectoryInfo subDir in dirs)
                {
                    string newDestinationDir = Path.Combine(destinationDir, subDir.Name);
                    CopyDirectory(subDir.FullName, newDestinationDir, true);
                }
            }
        }

    }
}
