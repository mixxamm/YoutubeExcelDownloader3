using ExcelDataReader;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Text.RegularExpressions;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.Storage;
using Windows.Storage.AccessCache;
using YoutubeExplode;
using YoutubeExplode.Converter;
using YoutubeExplode.Videos.Streams;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace YoutubeExcelDownloader3
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainWindow : Window
    {
        public MainWindow()
        {
            this.InitializeComponent();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        private void KiesBestandButton_Click(object sender, RoutedEventArgs e)
        {
            KiesBestandAsync();
        }

        private async void KiesBestandAsync()
        {
            var window = new Microsoft.UI.Xaml.Window();
            var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(window);
            var picker = new Windows.Storage.Pickers.FileOpenPicker();
            picker.ViewMode = Windows.Storage.Pickers.PickerViewMode.Thumbnail;
            picker.SuggestedStartLocation = Windows.Storage.Pickers.PickerLocationId.Downloads;
            picker.FileTypeFilter.Add(".xls");
            picker.FileTypeFilter.Add(".xlsx");
            WinRT.Interop.InitializeWithWindow.Initialize(picker, hwnd);
            Windows.Storage.StorageFile file = await picker.PickSingleFileAsync();
            if (file != null)
            {
                KiesBestandButton.Visibility = Visibility.Collapsed;
                // Application now has read/write access to the picked file
                // this.textBlock.Text = "Picked photo: " + file.Name;
                var stream = await file.OpenStreamForReadAsync();
                List<string> links = new List<string>();
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    TextBlockItems.Text = "Zoeken naar links...";
                    do
                    {
                        while (reader.Read())
                        {
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                string link = reader.GetString(i);
                                if (link != null && link.Contains("https://") && (link.Contains("youtube.com") || link.Contains("youtu.be")))
                                {
                                    links.Add(link);
                                }
                            }
                        }
                    } while (reader.NextResult());
                    TextBlockItems.Text = $"{links.Count} items gevonden.";
                    var folderPicker = new Windows.Storage.Pickers.FolderPicker();
                    folderPicker.SuggestedStartLocation = Windows.Storage.Pickers.PickerLocationId.Desktop;
                    folderPicker.FileTypeFilter.Add("*");
                    WinRT.Interop.InitializeWithWindow.Initialize(folderPicker, hwnd);
                    Windows.Storage.StorageFolder folder = await folderPicker.PickSingleFolderAsync();
                    if (folder != null)
                    {
                        // Application now has read/write access to all contents in the picked folder
                        // (including other sub-folder contents)
                        Windows.Storage.AccessCache.StorageApplicationPermissions.
                        FutureAccessList.AddOrReplace("PickedFolderToken", folder);
                        var youtube = new YoutubeClient();
                        TextBlockItems.Text = "Bezig met downloaden.";
                        int aantalLinks = links.Count;
                        int teller = 0;
                        int fouten = 0;
                        foreach (string link in links)
                        {
                            try
                            {
                                var video = await youtube.Videos.GetAsync(link);
                                var manifest = await youtube.Videos.Streams.GetManifestAsync(video.Id);
                                var streamInfo = manifest.GetAudioOnlyStreams().GetWithHighestBitrate();
                                if (streamInfo != null)
                                {
                                    var soundStream = await youtube.Videos.Streams.GetAsync(streamInfo);
                                    var storageFolder = await StorageApplicationPermissions.FutureAccessList.GetFolderAsync("PickedFolderToken");
                                    await youtube.Videos.DownloadAsync(link, $"{storageFolder.Path}/{Regex.Replace(video.Title, "[^a-zA-Z0-9_.]+", " ", RegexOptions.Compiled)}.mp3");
                                    teller++;
                                    DownloadProgress.Value = (teller + 0.0) / (aantalLinks + 0.0) * 100;
                                    TextBlockItems.Text = $"{teller - fouten}/{aantalLinks} liedjes gedownload." + (fouten > 0 ? $" Mislukt: {fouten}" : "");
                                }
                            }
                            catch (Exception e)
                            {
                                Debug.WriteLine(e);
                                teller++;
                                fouten++;
                            }
                        }
                    }
                    else
                    {
                        TextBlockItems.Text = "Downloaden geannuleerd.";
                    }

                }
            }
            else
            {
                // this.textBlock.Text = "Operation cancelled.";
            }
        }
    }
}
