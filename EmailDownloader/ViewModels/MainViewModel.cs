using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Controls;
using DevExpress.Mvvm;
using EmailDownloader.Models;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using S22.Imap;
using System.Windows;
using MahApps.Metro;

namespace EmailDownloader.ViewModels
{
    internal sealed class MainViewModel : ViewModelBase
    {
        private ImapClient imap;

        public MainViewModel()
        {
            InitDefaultValues();
        }

        private void InitDefaultValues()
        {
            Username = "@gmail.com";
            Hostname = "imap.gmail.com";
            Port = 993;
            Ssl = true;
            selectedConditionAdapter =
                new SearchConditionAdapter(SearchConditionGenerateMethods[SelectedConditionIndex]);
        }

        private void SetAppThemeByHostname()
        {
            string accent = "BaseLight";
            switch (hostname)
            {
                case "imap.gmail.com":
                    accent = "Red";
                    break;
                case "imap-mail.outlook.com":
                    accent = "Cobalt";
                    break;
                case "imap.zoho.com":
                    accent = "Steel";
                    break;
                case "imap.yandex.ru":
                    accent = "Red";
                    break;
            }
            ThemeManager.ChangeAppStyle(Application.Current,
              ThemeManager.GetAccent(accent), ThemeManager.GetAppTheme("BaseLight"));
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
                if (username != null && username.Equals(value))
                    return;

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
                if (hostname != null && hostname.Equals(value))
                    return;

                hostname = value;
                RaisePropertyChanged();

                SetAppThemeByHostname();
            }
        }

        public ObservableCollection<string> HostnameDefault => new ObservableCollection<string>
        {
            "imap.gmail.com",
            "imap-mail.outlook.com",
            "imap.zoho.com",
            "imap.yandex.ru"
        };

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
                     foundMessages = imap.Search(condition);
                     FoundMessagesCount = foundMessages.Count();
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
        private IEnumerable<uint> foundMessages;

        private int foundMessagesCount;
        public int FoundMessagesCount
        {
            get => foundMessagesCount;
            set
            {
                foundMessagesCount = value;
                RaisePropertyChanged();
            }
        }

        public AsyncCommand DownloadCommand => new AsyncCommand(() => DownloadAsync(), () => foundMessagesCount != 0 && Authed);
        private async Task DownloadAsync()
        {
            var progress = await ShowProgressAsync("Downloading", "Please wait...", true);
            int downloaded = 0;
            int errors = 0;
            try
            {
                string path = GenerateDirectoryForEmailMessages();
                await Task.Run(() =>
                {
                    foreach (var uid in foundMessages)
                    {
                        if (progress.IsCanceled)
                            break;

                        try
                        {
                            DownloadEmailMessage(downloaded, path, uid);
                        }
                        catch(Exception e)
                        {
                            Debug.Log(e);
                            errors++;
                            continue;
                        }
                        downloaded++;
                        ReportMessageDownloadProgress(progress, downloaded);
                    }
                });
                await ShowMessageAsync($"Downloading Complete ({downloaded}/{foundMessagesCount})",
                    $"Emails saved to {path}{Environment.NewLine}Failed count: {errors}");
            }
            catch(Exception e)
            {
                Debug.Log(e);
                await ShowMessageAsync($"Downloading Failed ({downloaded}/{foundMessagesCount})", $"{e.Message}");
            }
            await progress.CloseAsync();
        }

        private void DownloadEmailMessage(int downloaded, string path, uint uid)
        {
            using (var message = imap.GetMessage(uid))
            {
                string emailFolderPath = GenerateEmailMessageFolderPath(downloaded, path, message);

                SaveMessageHeaders(message, emailFolderPath);
                SaveMessageBody(message, emailFolderPath);
                SaveMessageAttachments(message, emailFolderPath);
                SaveMessageAlternativeViews(message, emailFolderPath);
            }
        }

        private void SaveMessageAlternativeViews(MailMessage message, string emailFolderPath)
        {
            if (message.AlternateViews.Count == 0)
                return;

            var avsPath = Directory.CreateDirectory(Path.Combine(emailFolderPath, "Alternative Views")).FullName;
            int id = 1;
            foreach (var av in message.AlternateViews)
            {
                var e = av.ContentType.MediaType.Split(new char[] { '/' })[1];
                var avPath = Path.Combine(avsPath, $"{id}.{e}");
                using (var stream = File.Create(avPath))
                {
                    av.ContentStream.CopyTo(stream);
                }
                id++;
            }
        }

        private void SaveMessageHeaders(MailMessage message, string emailFolderPath)
        {
            var headersPath = Path.Combine(emailFolderPath, "Headers.txt");
            File.Create(headersPath).Close();

            var headersLines = new List<string>();
            for(int i = 0; i < message.Headers.Count; i++)
            {
                headersLines.Add($"{message.Headers.Keys[i]}: {message.Headers.GetValues(i).First()}");
            }
            File.WriteAllLines(headersPath, headersLines);
        }

        private void SaveMessageBody(MailMessage message, string emailFolderPath)
        {
            string body = message.Body;
            var bodyPath = Path.Combine(emailFolderPath, "Body.html");

            File.Create(bodyPath).Close();
            File.WriteAllText(bodyPath, body);
        }

        private void SaveMessageAttachments(MailMessage message, string emailFolderPath)
        {
            if (message.Attachments.Count == 0)
                return;

            var attachmentsPath = Directory.CreateDirectory(Path.Combine(emailFolderPath, "Attachments")).FullName;
            int id = 1;
            foreach (var attachment in message.Attachments)
            {
                var attachmentPath = Path.Combine(attachmentsPath, $"{id}{Path.GetExtension(attachment.Name)}");
                using (var stream = File.Create(attachmentPath))
                {
                    attachment.ContentStream.CopyTo(stream);
                }
                id++;
            }
        }

        private void ReportMessageDownloadProgress(ProgressDialogController progress, int downloaded)
        {
            progress.SetProgress((double)downloaded / foundMessagesCount);
            progress.SetMessage($"Please wait...({downloaded}/{foundMessagesCount})");
        }

        private string GenerateEmailMessageFolderPath(int downloaded, string path, MailMessage message)
        {
            var emailFolderName = $"{downloaded}_{message.Subject}";
            emailFolderName = $"{new string(emailFolderName.Take(30).ToArray())}";
            foreach (char c in Path.GetInvalidFileNameChars())
            emailFolderName = emailFolderName.Replace(c, '_');
            var emailFolderPath = Path.Combine(path, emailFolderName);
            return Directory.CreateDirectory(emailFolderPath).FullName;
        }

        private string GenerateDirectoryForEmailMessages()
        {
            var path = Path.Combine(Environment.CurrentDirectory, $"{username}_{DateTime.Now.ToFileTime()}");
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
