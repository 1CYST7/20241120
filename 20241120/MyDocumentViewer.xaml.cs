using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace _20241120
{
    /// <summary>
    /// MyDocumentViewer.xaml 的互動邏輯
    /// </summary>
    public partial class MyDocumentViewer : Window
    {
        Color fontColor = Colors.Black;
        
        public MyDocumentViewer()
        {
            InitializeComponent();
            fontColorPicker.SelectedColor = fontColor;
            // 匯入字體
            foreach (FontFamily fontFamily in Fonts.SystemFontFamilies)
            {
                fontFamilyComboBox.Items.Add(fontFamily.Source);
            }
            //設置預設字體為 "微軟正黑體"
            fontFamilyComboBox.SelectedIndex = 8;

            fontSizeComboBox.ItemsSource = new List<double>() { 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 };
            // 設置預設字體大小為 10(第3個位置)
            fontSizeComboBox.SelectedIndex = 2;

            // 設置 BackgroundColorPicker 的預設選擇顏色為黑色
            BackgroundColorPicker.SelectedColor = Colors.Black;
        }

        private void NewCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            //建立新的 MyDocumentViewer 並顯示
            MyDocumentViewer myDocumentViewer = new MyDocumentViewer();
            myDocumentViewer.Show();
        }

        private void OpenCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {

            // 創建一個新的打開文件對話框
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // 設置文件過濾器，允許打開 RTF 和 HTML 文件，以及所有文件
            openFileDialog.Filter = "Rich Text Format (*.rtf)|*.rtf|HTML Files (*.html;*.htm)|*.html;*.htm|All files (*.*)|*.*";

            // 設置默認文件擴展名為 .rtf
            openFileDialog.DefaultExt = ".rtf";

            // 添加擴展名到文件名中
            openFileDialog.AddExtension = true;

            if (openFileDialog.ShowDialog() == true)
            {
                // 創建一個文件流來打開選擇的文件
                FileStream fileStream = new FileStream(openFileDialog.FileName, FileMode.Open);

                // 創建一個 TextRange 來表示 RichTextBox 中的所有文本
                TextRange range = new TextRange(documentRichTextBox.Document.ContentStart, documentRichTextBox.Document.ContentEnd);

                if (openFileDialog.FileName.EndsWith(".html") || openFileDialog.FileName.EndsWith(".htm"))
                {
                    // 使用 StreamReader 來讀取文件內容
                    using (StreamReader reader = new StreamReader(fileStream))
                    {
                        // 讀取 HTML 文件的所有內容
                        string htmlContent = reader.ReadToEnd();

                        // 將 HTML 內容轉換為純文本並設置到 TextRange 中
                        range.Text = HtmlToText(htmlContent);
                    }
                }else
                {
                    //將文件內容加載為 RTF 格式
                    range.Load(fileStream, DataFormats.Rtf);
                }

                fileStream.Close();
            }
        }
        private string HtmlToText(string html)
        {
            // 這裡可以使用第三方庫如 HtmlAgilityPack 來解析 HTML
            // 簡單的實現可以使用正則表達式來去除 HTML 標籤
            return System.Text.RegularExpressions.Regex.Replace(html, "<.*?>", string.Empty);
        }

        private void SaveCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            // 創建一個新的保存文件對話框
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            // 設置文件過濾器，允許保存為 RTF 和 HTML 文件，以及所有文件
            saveFileDialog.Filter = "Rich Text Format (*.rtf)|*.rtf|HTML Files (*.html;*.htm)|*.html;*.htm|All files (*.*)|*.*";

            // 設置默認文件擴展名為 .rtf
            saveFileDialog.DefaultExt = ".rtf";

            // 添加擴展名到文件名中
            saveFileDialog.AddExtension = true;

            // 是否在保存文件對話框中選擇了一個文件並點擊了確定
            if (saveFileDialog.ShowDialog() == true)
            {
                // 創建一個 TextRange 來表示 RichTextBox 中的所有文本
                TextRange range = new TextRange(documentRichTextBox.Document.ContentStart, documentRichTextBox.Document.ContentEnd);

                // 創建一個文件流來創建選擇的文件
                FileStream fileStream = new FileStream(saveFileDialog.FileName, FileMode.Create);

                if (saveFileDialog.FileName.EndsWith(".html") || saveFileDialog.FileName.EndsWith(".htm"))
                {
                    // 將文本轉換為 HTML 格式
                    string htmlContent = TextToHtml(range.Text);

                    // 使用 StreamWriter 將 HTML 內容寫入文件
                    using (StreamWriter writer = new StreamWriter(fileStream))
                    {
                        writer.Write(htmlContent);
                    }
                }
                else{
                    // 否則，將文本內容保存為 RTF 格式
                    range.Save(fileStream, DataFormats.Rtf);
                }
                fileStream.Close();
            }
        }
        private string TextToHtml(string text)
        {
            // 這裡可以使用第三方庫如 HtmlAgilityPack 來生成 HTML
            // 簡單的實現可以使用基本的 HTML 標籤來包裹文本
            return $"<html><body>{System.Net.WebUtility.HtmlEncode(text).Replace("\n", "<br>")}</body></html>";
        }
        private void documentRichTextBox_SelectionChanged(object sender, RoutedEventArgs e)
        {
            // 獲取當前選擇文本的字體粗細屬性
            var property_bold = documentRichTextBox.Selection.GetPropertyValue(TextElement.FontWeightProperty);

            // 設置 boldToggleButton 的選中狀態，如果選擇的文本是粗體，則設置為選中狀態
            boldToggleButton.IsChecked = (property_bold != DependencyProperty.UnsetValue) && (property_bold.Equals(FontWeights.Bold));

            // 獲取當前選擇文本的字體樣式屬性
            var property_italic = documentRichTextBox.Selection.GetPropertyValue(TextElement.FontStyleProperty);

            // 設置 italicToggleButton 的選中狀態，如果選擇的文本是斜體，則設置為選中狀態
            italicToggleButton.IsChecked = (property_italic != DependencyProperty.UnsetValue) && (property_italic.Equals(FontStyles.Italic));

            // 獲取當前選擇文本的下劃線屬性
            var property_underline = documentRichTextBox.Selection.GetPropertyValue(Inline.TextDecorationsProperty);

            // 設置 underlineToggleButton 的選中狀態，如果選擇的文本有下劃線，則設置為選中狀態
            underlineToggleButton.IsChecked = (property_underline != DependencyProperty.UnsetValue) && (property_underline.Equals(TextDecorations.Underline));

            // 獲取當前選擇文本的字體大小屬性
            var property_fontSize = documentRichTextBox.Selection.GetPropertyValue(TextElement.FontSizeProperty);

            // 設置 fontSizeComboBox 的選中項目為當前選擇文本的字體大小
            fontSizeComboBox.SelectedItem = property_fontSize;

            // 獲取當前選擇文本的字體家族屬性
            var property_fontFamily = documentRichTextBox.Selection.GetPropertyValue(TextElement.FontFamilyProperty);

            // 設置 fontFamilyComboBox 的選中項目為當前選擇文本的字體家族
            fontFamilyComboBox.SelectedItem = property_fontFamily.ToString();

            // 獲取當前選擇文本的前景色屬性
            var property_fontColor = documentRichTextBox.Selection.GetPropertyValue(TextElement.ForegroundProperty);

            // 如果前景色屬性是 SolidColorBrush，則設置 fontColorPicker 的選中顏色為該顏色
            if (property_fontColor is SolidColorBrush solidColorBrush)
            {
                fontColorPicker.SelectedColor = solidColorBrush.Color;
            }else{
                // 如果前景色屬性不是 SolidColorBrush，則設置 fontColorPicker 的選中顏色為透明或其他默認顏色
                fontColorPicker.SelectedColor = Colors.Transparent; // 或其他默認顏色
            }
        }

        private void fontSizeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (fontSizeComboBox.SelectedItem != null)
            {
                documentRichTextBox.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, fontSizeComboBox.SelectedItem);
            }
        }

        private void fontFamilyComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (fontFamilyComboBox.SelectedItem != null)
            {
                documentRichTextBox.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, new FontFamily((string)fontFamilyComboBox.SelectedItem));
            }
        }

        private void fontColorPicker_SelectedColorChanged(object sender, RoutedPropertyChangedEventArgs<Color?> e)
        {
            fontColor = fontColorPicker.SelectedColor.Value;
            SolidColorBrush brush = new SolidColorBrush(fontColor);
            // 將 RichTextBox 中當前選擇的文本的前景色屬性設置為指定的 SolidColorBrush
            documentRichTextBox.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, brush);
        }

        private void clearButton_Click(object sender, RoutedEventArgs e)
        {
            documentRichTextBox.Document.Blocks.Clear();
        }

        private void BackgroundColorPicker_SelectedColorChanged(object sender, RoutedPropertyChangedEventArgs<Color?> e)
        {
            // 檢查新顏色是否有值
            if (e.NewValue.HasValue){
                // 獲取選擇的顏色
                Color selectedColor = e.NewValue.Value;
                // 創建一個新的 SolidColorBrush，使用選擇的顏色
                SolidColorBrush brush = new SolidColorBrush(selectedColor);
                // 將 RichTextBox 中當前選擇的文本的背景色屬性設置為指定的 SolidColorBrush
                documentRichTextBox.Selection.ApplyPropertyValue(TextElement.BackgroundProperty, brush);
            }
        }
    }
}
