using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileWatcherService
{
    class FileWatcher
    {
        FileSystemWatcher watcher;

        public string MonitorDirectory
        {
            get { return watcher != null ? watcher.Path : null; }
        }

        public string FileFilter
        {
            get { return watcher != null ? watcher.Filter : null; }
        }

        public FileWatcher(string monitorDirectory, string fileFilter)
        {
            watcher = new FileSystemWatcher(monitorDirectory, fileFilter);
            SetFileWatcherFilters();
            AddFileWatcherEventHandlers();
        }

        void AddFileWatcherEventHandlers()
        {
            if (watcher == null)
                return;
            watcher.Created += Watcher_Created;
            watcher.Changed += Watcher_Changed;
            watcher.Error += Watcher_Error;
        }

        void SetFileWatcherFilters()
        {
            if (watcher == null)
                return;
            watcher.NotifyFilter = NotifyFilters.Attributes |
                NotifyFilters.CreationTime |
                NotifyFilters.DirectoryName |
                NotifyFilters.FileName |
                NotifyFilters.LastWrite |
                NotifyFilters.Size;
        }

        private void Watcher_Error(object sender, ErrorEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void Watcher_Changed(object sender, FileSystemEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void Watcher_Created(object sender, FileSystemEventArgs e)
        {
            //TODO: 
            //Check file name
            //If a valid file name send it to FileReader
            //Create security objects
            //Send to DAL directly or through WCF??
        }
    }
}
