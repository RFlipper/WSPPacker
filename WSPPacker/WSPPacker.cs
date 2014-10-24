using Microsoft.Deployment.Compression;
using Microsoft.Deployment.Compression.Cab;
using OY.TotalCommander.TcPluginInterface;
using OY.TotalCommander.TcPluginInterface.Packer;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;

namespace Plumsail.WSPPacker
{
    public class WSPPacker : PackerPlugin
    {
        public const string AllowedExtensionsOnForceShow = ".WSP,.CAB";

        /// <summary>
        /// Wrap pack/unpack functionality
        /// </summary>
        private class ArchiveInfo : IDisposable
        {
            public readonly string OriginalPath;
            public readonly string TempFolder = Path.GetTempPath() + Guid.NewGuid();

            public ArchiveInfo(string path)
            {
                OriginalPath = path;

                Directory.CreateDirectory(TempFolder);
                if (File.Exists(OriginalPath))
                {
                    var cab = new CabInfo(OriginalPath);
                    cab.Unpack(TempFolder);
                }
            }
            public void Dispose()
            {
                var cab = new CabInfo(OriginalPath);
                cab.Pack(TempFolder, true, CompressionLevel.Normal, (sender, args) => { });

                Directory.Delete(TempFolder, true);
            }
        }

        private IEnumerator _archEnumerator;

        public WSPPacker(StringDictionary pluginSettings) : base(pluginSettings)
        {
            Capabilities = PackerCapabilities.New | PackerCapabilities.Modify
                | PackerCapabilities.Multiple | PackerCapabilities.Delete;
        }

        #region IPackerPlugin Members

        /// <summary>
        /// Prepare archive to open 
        /// </summary>
        /// <param name="archiveData"></param>
        /// <returns></returns>
        public override object OpenArchive(ref OpenArchiveData archiveData)
        {
            var archive        = new CabInfo(archiveData.ArchiveName);
            _archEnumerator    = archive.GetFiles().GetEnumerator();
            archiveData.Result = PackerResult.OK;

            return archive;
        }

        /// <summary>
        /// Read information about files in archieve
        /// </summary>
        /// <param name="arcData"></param>
        /// <param name="headerData"></param>
        /// <returns></returns>
        public override PackerResult ReadHeader(ref object arcData, out HeaderData headerData)
        {
            headerData = null;
            CabInfo cf = arcData as CabInfo;

            var pathLength = cf.ToString().Length;
            if (_archEnumerator.MoveNext())
            {
                object current = _archEnumerator.Current;

                if (current is CabFileInfo)
                {
                    headerData = new HeaderData();
                    GetHeaderData((CabFileInfo)current, ref headerData, pathLength);
                }
                else if (current != null)
                    throw new InvalidOperationException("Unknown type in FindNext: " + current.GetType().FullName);

                return PackerResult.OK;
            }
            else
            {
                return PackerResult.EndArchive;
            }
        }
        
        /// <summary>
        /// Fill file information
        /// </summary>
        /// <param name="entry"></param>
        /// <param name="headerData"></param>
        /// <param name="padLeft"></param>
        private static void GetHeaderData(CabFileInfo entry, ref HeaderData headerData, int padLeft)
        {
            headerData.FileName       = entry.FullName.Remove(0, padLeft);
            headerData.PackedSize     = (ulong)entry.Length;
            headerData.UnpackedSize   = (ulong)entry.Length;
            headerData.FileTime       = entry.CreationTime;
            headerData.FileAttributes = entry.Attributes;
            headerData.ArchiveName    = entry.ArchiveName;
        }

        /// <summary>
        /// Process single file
        /// </summary>
        /// <param name="arcData"></param>
        /// <param name="operation"></param>
        /// <param name="destFile"></param>
        /// <returns></returns>
        public override PackerResult ProcessFile(object arcData, ProcessFileOperation operation, string destFile)
        {
            switch (operation)
            {
                case ProcessFileOperation.Extract:
                    CabInfo cf   = arcData as CabInfo;
                    var file     = _archEnumerator.Current as CabFileInfo;
                    var filePath = file.FullName.Remove(0, cf.ToString().Length + 1);

                    cf.UnpackFile(filePath, destFile);

                    return PackerResult.OK;

                default:
                    return PackerResult.OK;
            }
        }

        /// <summary>
        /// Pack files in archieve
        /// </summary>
        /// <param name="packedFile"></param>
        /// <param name="subPath"></param>
        /// <param name="srcPath"></param>
        /// <param name="addList"></param>
        /// <param name="flags"></param>
        /// <returns></returns>
        public override PackerResult PackFiles(string packedFile, string subPath, string srcPath, List<string> addList, PackFilesFlags flags)
        {
            using (var archieve = new ArchiveInfo(packedFile))
            {
                foreach (var file in addList)
                {
                    var from = Path.Combine(srcPath, file);
                    var to   = archieve.TempFolder;

                    if (!string.IsNullOrEmpty(subPath))
                        to = Path.Combine(archieve.TempFolder, subPath);

                    to = Path.Combine(to, file);

                    FileAttributes attr = File.GetAttributes(from);
                    if ((attr & FileAttributes.Directory) == FileAttributes.Directory)
                    {
                        Directory.CreateDirectory(to);
                    }
                    else
                    {
                        File.Copy(from, to, true);
                    }
                }
            }
        
            return PackerResult.OK;
        }

        /// <summary>
        /// Delete files from archieve
        /// </summary>
        /// <param name="packedFile"></param>
        /// <param name="deleteList"></param>
        /// <returns></returns>
        public override PackerResult DeleteFiles(string packedFile, List<string> deleteList)
        {
            using (var archieve = new ArchiveInfo(packedFile))
            {
                foreach (var file in deleteList)
                {
                    var path = Path.Combine(archieve.TempFolder, file.TrimStart(new char[] {'\\'}));
                    System.IO.File.Delete(path);
                }
            }

            return PackerResult.OK;
        }

        public override void ConfigurePacker(TcWindow parentWin)
        {
            System.Windows.Forms.MessageBox.Show(parentWin,
                "Provides WSP/CAB packer functionality\n" +
                "http://plumsail.com\n\n" +
                "Author: Rylov Roman",
                "WSP/CAB Packer PlugIn",
                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        }

        public override bool CanYouHandleThisFile(string fileName)
        {
            string extension = Path.GetExtension(fileName);
            return extension != null && AllowedExtensionsOnForceShow.Contains(extension.ToUpper());
        }
        public override PackerResult CloseArchive(object arcData)
        {
            return PackerResult.OK;
        }

        #region Optional Methods

        #endregion Optional Methods

        #endregion IPackerPlugin Members
    }
}
