using System;
using System.Diagnostics;
using System.IO;
using Microsoft.Deployment.Compression;
using Microsoft.Deployment.Compression.Cab;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OY.TotalCommander.TcPluginInterface.Packer;
using System.Runtime.InteropServices;

namespace WSPPacker.Tests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void GetFileTest()
        {
            var arc = new CabInfo(@"c:\Temp\TotalCmd\CT-Mike2.wsp");
            var files = arc.GetFiles();
            Debug.Assert(files.Count > 1 );
        }

        [TestMethod]
        public void ExtractFileTest()
        {
            var arc = new CabInfo(@"c:\Temp\TotalCmd\CT-Mike2.wsp");
            arc.UnpackFile(@"CT-Mike2Workflows\Elements.xml", @"c:\Temp\1\Elements.xml");
        }

        [TestMethod]
        public void ExtractAllFilesTest()
        {
            var arc = new CabInfo(@"c:\Temp\TotalCmd\1.wsp");
            arc.Unpack(@"c:\Temp\TotalCmd\Test");
        }


        [TestMethod]
        public void PackFileTest()
        {
            var arc = new CabInfo(@"c:\Temp\TotalCmd\2.wsp");
            arc.Pack(@"c:\Users\RFLIP_~1\AppData\Local\Temp\2813b241-5632-4494-8e9b-faa40dcb3cfa", true,CompressionLevel.Normal, (sender, args) => {} );
        }


        [TestMethod]
        public void InitArch()
        {
            var arc = new CabInfo(@"c:\Temp\TotalCmd\CT-Mike2.wsp");
            var packer = new Plumsail.WSPPacker.WSPPacker(null);

            IntPtr ptr = Marshal.AllocHGlobal(4);
//Marshal.WriteInt32(ptr, value);
            var OAData = new OpenArchiveData(ptr, false);
            var archive = packer.OpenArchive(ref OAData);
            var header = new HeaderData();
            packer.ReadHeader(ref archive, out header);
        }
    }
}
