using System;
using System.IO;

namespace FilesToBytesETC
{
    class Program
    {
        public static byte[] ConvertFileToBytes()
        {
            var path = @"D:\Excel - Copy.xls";
            //var path = Directory.GetFiles(path1);
            FileStream fs = new FileStream(path, FileMode.Open,FileAccess.Read);
            BinaryReader br = new BinaryReader(fs);
            long numbytes = new FileInfo(path).Length;
            byte[] buffer = br.ReadBytes((int)numbytes);
            return buffer;
        }
        static void Main(string[] args)
        {
            var b = ConvertFileToBytes();
            foreach (var item in b)
            {
                Console.Write(item);
            }
            Console.WriteLine("File byte created.");
        }
    }
}
