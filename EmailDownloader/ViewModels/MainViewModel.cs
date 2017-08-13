using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Threading.Tasks;
using DevExpress.Mvvm;
using S22.Imap;
using System.Windows.Controls;
using EmailDownloader.Models;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using System.Collections.ObjectModel;
using System.Net.Mail;
using System.IO;
using System.Diagnostics;

namespace EmailDownloader.ViewModels
{
    internal sealed class MainViewModel : ViewModelBase
    {
        private ImapClient imap;
        private string savePattern;

        public MainViewModel()
        {
            savePattern = File.ReadAllText("savePattern.txt");
            InitialDefaultValues();
        }
        private void InitialDefaultValues()
        {
            Username = "@gmail.com";
            Hostname = "imap.gmail.com";
            Port = 993;
            Ssl = true;

            selectedConditionAdapter = new SearchConditionAdapter(SearchConditionGenerateMethods[SelectedConditionIndex]);
        }

        #region Authorization
        public bool Authed => imap != null && imap.Authed;
        public bool MayAuth => !Authed && !string.IsNullOrWhiteSpace(username) && !string.IsNullOrWhiteSpace(hostname) && !AuthCommand.IsExecuting;

        private string username;
        public string Username
        {
            get => username;
            set
            {
                username = value;
                RaisePropertyChanged();
            }
        }

        private string hostname;
        public string Hostname
        {
            get => hostname;
            set
            {
                hostname = value;
                RaisePropertyChanged();
            }
        }

        private int port;
        public int Port
        {
            get => port;
            set
            {
                port = value;
                RaisePropertyChanged();
            }
        }

        private bool ssl;
        public bool Ssl
        {
            get => ssl;
            set
            {
                ssl = value;
                RaisePropertyChanged();
            }
        }

        public AsyncCommand<PasswordBox> AuthCommand => 
            new AsyncCommand<PasswordBox>((PasswordBox passwordPanel) => AuthAsync(passwordPanel.Password),
            (PasswordBox passwordPanel) => MayAuth);
        private async Task AuthAsync(string password)
        {
            var progress = await ShowProgressAsync("Authentication", "Please wait...");
            progress.SetIndeterminate();
            try
            {
                if (await Task.Run(() => !InitializeImapTask(hostname, port, username, password, AuthMethod.Auto, ssl).Wait(6000)))
                {
                    await AuthTimeout();
                }
            }
            catch (Exception e)
            {
                if (e.InnerException != null)
                    await AuthFailed(e.InnerException.Message);
                else
                    await AuthFailed(e.Message);
            }
            await AuthSucced();

            async Task AuthSucced()
            {
                await progress.CloseAsync();
            }
            async Task AuthFailed(string cause)
            {
                await progress.CloseAsync();
                await ShowMessageAsync("Authentication Failed", cause);
            }
            async Task AuthTimeout()
            {
                await progress.CloseAsync();
                await ShowMessageAsync("Authentication Timeout", "Six seconds passed but still no response");
            }
        }
        private Task InitializeImapTask(string hostname, int port,
            string username, string password, AuthMethod auth, bool ssl)
        {
            return Task.Factory.StartNew(() => imap = new ImapClient(hostname, port, username, password, AuthMethod.Auto, ssl));
        }

        public DelegateCommand DisconnectCommand => new DelegateCommand(() =>
        {
            imap.Dispose();
            imap = null;
        }, () => Authed && !SearchCommand.IsExecuting);
        #endregion

        #region Search
        private int selectedConditionIndex;
        public int SelectedConditionIndex
        {
            get => selectedConditionIndex;
            set
            {
                if (selectedConditionIndex == value)
                    return;

                selectedConditionIndex = value;
                RaisePropertyChanged();

                SelectedConditionAdapter = new SearchConditionAdapter(SearchConditionGenerateMethods[SelectedConditionIndex]);
            }
        }

        private SearchConditionAdapter selectedConditionAdapter;
        public SearchConditionAdapter SelectedConditionAdapter
        {
            get => selectedConditionAdapter;
            set
            {
                if (selectedConditionAdapter != null && selectedConditionAdapter.Equals(value))
                    return;

                selectedConditionAdapter = value;
                RaisePropertyChanged();
            }
        }

        public ObservableCollection<MethodInfo> SearchConditionGenerateMethods =>
            new ObservableCollection<MethodInfo>(typeof(SearchCondition).GetMethods().Where((m) => 
            {
                bool result = false;
                result = m.ReturnType.Equals(typeof(SearchCondition));
                if(m.GetParameters().Count() != 0)
                {
                    if (m.GetParameters().First().ParameterType.Equals(typeof(SearchCondition)))
                        result = false;
                    else if (m.GetParameters().First().ParameterType.Equals(typeof(UInt32[])))
                        result = false;
                }
                return result;
            }));

        public AsyncCommand SearchCommand => new AsyncCommand(() => SearchAsync(), () => Authed && !SearchCommand.IsExecuting);
        private async Task SearchAsync()
        {
            var progress = await ShowProgressAsync("Searching", "Please wait...");
            progress.SetIndeterminate();
            try
            {
                await Task.Run(() =>
                 {
                     var condition = selectedConditionAdapter.Method.Invoke(null, selectedConditionAdapter.GetValues()) as SearchCondition;
                     messagesToDownload = imap.Search(condition);
                     MessagesToDownloadCount = messagesToDownload.Count();
                 });
            }
            catch(Exception e)
            {
                await ShowMessageAsync("Search Failed", e.Message);
            }
            await progress.CloseAsync();
        }
        #endregion

        #region Download
        private IEnumerable<uint> messagesToDownload;

        private int messagesToDownloadCount;

        public int MessagesToDownloadCount
        {
            get => messagesToDownloadCount;
            set
            {
                messagesToDownloadCount = value;
                RaisePropertyChanged();
            }
        }

        public AsyncCommand DownloadCommand => new AsyncCommand(() => DownloadAsync(), () => messagesToDownloadCount != 0 && Authed);
        private async Task DownloadAsync()
        {
            ProgressDialogController progress = await ShowProgressAsync("Downloading", "Please wait...", true);
            int downloaded = 0;
            int errors = 0;
            try
            {
                string path = NewPath();
                await Task.Run(() =>
                {
                    foreach (var uid in messagesToDownload)
                    {
                        if (progress.IsCanceled)
                            break;

                        try
                        {
                            var message = imap.GetMessage(uid);
                            string filePath = GenerateFilePath(downloaded, path, message);
                            File.WriteAllText(filePath,
                                string.Format(savePattern, 
                                new object[] { message.From.Address, message.To.First(),
                                    message.Date(), message.Subject, message.Body }));
                        }
                        catch
                        {
                            errors++;
                            continue;
                        }
                        downloaded = ReportProgress(progress, downloaded);
                    }
                });
                await ShowMessageAsync($"Downloading Complete ({downloaded}/{messagesToDownloadCount})",
                    $"Emails saved to {path}{Environment.NewLine}Failed count: {errors}");
            }
            catch(Exception e)
            {
                await ShowMessageAsync($"Downloading Failed ({downloaded}/{messagesToDownloadCount})", $"{e.Message}");
            }
            await progress.CloseAsync();
        }

        private int ReportProgress(ProgressDialogController progress, int downloaded)
        {
            downloaded++;
            progress.SetProgress((double)downloaded / messagesToDownloadCount);
            progress.SetMessage($"Please wait...({downloaded}/{messagesToDownloadCount})");
            return downloaded;
        }

        private static string GenerateFilePath(int downloaded, string path, MailMessage message)
        {
            var fileName = $"{downloaded}_{message.Subject}";
            fileName = $"{new string(fileName.Take(30).ToArray())}.html";
            foreach (char c in Path.GetInvalidFileNameChars())
                fileName = fileName.Replace(c, '_');
            var filePath = Path.Combine(path, fileName);
            return filePath;
        }

        private string NewPath()
        {
            var path = Path.Combine(Environment.CurrentDirectory, $"{username}_{hostname}_{DateTime.Now.ToFileTime()}");
            foreach (char c in Path.GetInvalidPathChars())
                path = path.Replace(c, '_');
            Directory.CreateDirectory(path);
            return path;
        }
        #endregion

        #region Helpers
        private async Task<ProgressDialogController> ShowProgressAsync(string title, string message, bool isCancelable = false)
        {
            var mainWindow = WindowAccessor.Main as MetroWindow;
            return await mainWindow.Invoke(() => mainWindow.ShowProgressAsync(title, message, isCancelable));
        }
        private async Task ShowMessageAsync(string title, string message)
        {
            var mainWindow = WindowAccessor.Main as MetroWindow;
            await mainWindow.Invoke(() => mainWindow.ShowMessageAsync(title, message));
        }
        #endregion
    }
}
