using System;
using System.Collections.Generic;
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
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Threading;
using System.Reflection;
using System.Drawing;

namespace AppDevReportGenerator
{
    /// <summary>
    /// Interaction logic for ReportWindow.xaml
    /// </summary>
    public partial class ReportWindow : Window
    {
        private Report ActiveReport = new Report();
        private readonly SynchronizationContext SynchronizationContextObject = SynchronizationContext.Current;
        private string ReportsDirectory = System.IO.Path.Combine(Environment.CurrentDirectory, "JSON");

        public ReportWindow()
        {
            InitializeComponent();
            ContentRendered += ReportWindow_ContentRendered;
            SizeChanged += ReportWindow_SizeChanged;
            ToggleExportButton(false);
        }

        private void ReportWindow_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            ReportBorder.Height = this.ActualHeight - 180;
            AllBorder.Height = this.ActualHeight - 180;
            SelectedBorder.Height = this.ActualHeight - 180;
        }

        private void ReportWindow_ContentRendered(object sender, EventArgs e)
        {
            LoadReports();
        }
        
        public void LoadReports()
        {
            LoadingRect.Visibility = Visibility.Visible;
            Loading.Visibility = Visibility.Visible;
            Mouse.OverrideCursor = Cursors.AppStarting;

            ReportsPanel.Children.Clear();
            AllFieldsPanel.Children.Clear();
            SelectedFieldsPanel.Children.Clear();

            DirectoryInfo reportsdirectoryinfo = new DirectoryInfo(ReportsDirectory);
            foreach(FileInfo reportfileinfo in reportsdirectoryinfo.GetFiles("*.json"))
            {
                Report report = new Report(reportfileinfo.FullName);
                ReportsPanel.Children.Add(GetReportButton(report));
            }

            LoadingRect.Visibility = Visibility.Collapsed;
            Loading.Visibility = Visibility.Collapsed;
            Mouse.OverrideCursor = Cursors.Arrow;
        }

        private Border GetReportButton(Report report)
        {
            Button button = new Button {
                Content = report.Name,
                Style = Resources["Flat_Button"] as Style,
                Tag = report,
                UseLayoutRounding =true
            };
            button.MouseEnter += Button_MouseEnter;
            button.MouseLeave += Button_MouseLeave;
            button.Click += ReportButton_Clicked;

            Border border = new Border {
                BorderBrush = System.Windows.Media.Brushes.DimGray,
                BorderThickness = new Thickness(2),
                CornerRadius = new CornerRadius(5),
                Margin = new Thickness(5),
                UseLayoutRounding = true,
                Child = button
            };
            return border;
        }

        private void Button_MouseLeave(object sender, MouseEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Arrow;
        }

        private void Button_MouseEnter(object sender, MouseEventArgs e)
        {
            if (((Button)sender).IsEnabled) { Mouse.OverrideCursor = Cursors.Hand; }
            else { Mouse.OverrideCursor = Cursors.No; }
        }

        private void ReportButton_Clicked(object sender, RoutedEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Wait;
            LoadReports();
            try
            {
                AllFieldsPanel.Children.Clear();
                SelectedFieldsPanel.Children.Clear();
                Button selectedbutton = (Button)sender;
                ActiveReport = selectedbutton.Tag as Report;
                Header.Content = $"AppDev Report Generator: {ActiveReport.Name}";

                #region add to panels - fixed for worksheet columns in wrong order
                ActiveReport.GetReportHeaders();

                //ServicePro exports in a random order, so must get column id by name
                int sourceindex = 1;
                foreach (string header in ActiveReport.DefinitionHeaders)
                {
                    ReportField field = ActiveReport.Fields.Where(x => x.ExportName.ToLower() == header.ToLower()).First();
                    field.SourceIndex = sourceindex;
                    sourceindex++;
                }
                foreach (ReportField field in ActiveReport.Fields)
                {
                    Border fieldbutton = GetReportFieldButton(field);
                    if (field.ExportIndex > 0) { SelectedFieldsPanel.Children.Add(fieldbutton); }  // Automatically add fields with an export index (in the JSON file) greater than 0 to the right panel
                    else { AllFieldsPanel.Children.Add(fieldbutton); }  // Fields that are defined but have an export value less than or equal to zero go in the middle panel
                }
                #endregion

                ToggleExportButton(true);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Unable to load report definition: {ex.StackTrace}");
                MessageBox.Show($"Unable to load report definition: {ex.Message}.", "Error Loading Report");
            }
            finally
            {
                Mouse.OverrideCursor = Cursors.Arrow;
            }
        }

        private void ToggleExportButton(bool enable)
        {
            if(enable)
            {
                ExportButton.IsEnabled = true;
                ExportButton.Opacity = 1;
                ExportButton.Foreground = System.Windows.Media.Brushes.Black;
            }
            else
            {
                ExportButton.IsEnabled = false;
                ExportButton.Opacity = .6;
                ExportButton.Foreground = System.Windows.Media.Brushes.Gray;
            }
        }

        private Border GetReportFieldButton(ReportField field)
        {
            Button button = new Button
            {
                Content = field.Name,
                Style = Resources["Flat_Button"] as Style,
                UseLayoutRounding = true
            };
            button.MouseEnter += Button_MouseEnter;
            button.MouseLeave += Button_MouseLeave;

            Border border = new Border
            {
                BorderBrush = System.Windows.Media.Brushes.DimGray,
                BorderThickness = new Thickness(2),
                CornerRadius = new CornerRadius(5),
                Margin = new Thickness(5),
                Tag = field,
                UseLayoutRounding = true,
                Child = button
            };
            border.Drop += Border_Drop;
            border.PreviewMouseMove += Border_MouseMove;
            return border;
        }

        private void Border_MouseMove(object sender, MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                DataObject data = new DataObject();
                data.SetData(((Border)sender));
                DragDrop.DoDragDrop(this, data, DragDropEffects.Move);
            }
        }

        private void Border_Drop(object sender, DragEventArgs e)
        {
            try
            {
                Border border = (Border)e.Data.GetData(typeof(Border));
                System.Windows.Point dropped_at = e.GetPosition(this);
                StackPanel source_panel = GetParent(border) as StackPanel;
                DependencyObject hit_source = this.InputHitTest(dropped_at) as DependencyObject;
                StackPanel target_panel = GetParent(hit_source) as StackPanel;    
                UIElement dropped_on = sender as Border;

                if (target_panel != null)
                {
                    e.Effects = DragDropEffects.Move;
                    source_panel.Children.Remove(border);
                    int index = target_panel.Children.IndexOf(dropped_on) <  0 ? 0 : target_panel.Children.IndexOf(dropped_on);
                    target_panel.Children.Insert(index, border);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error dropping on border {ex.StackTrace}");
                MessageBox.Show($"Error dropping border: {ex.Message}.", "Border Drop Error");
            }
            finally
            {
                e.Handled = true;
            }
        }

        private DependencyObject GetParent(DependencyObject control, int recursioncount = 0)
        {
            DependencyObject parent = VisualTreeHelper.GetParent(control);
            while(parent.GetType() != typeof(StackPanel) && recursioncount < 10)  // Not the right way to do this, but sometimes you have to crawl up a few (unknown at compile time) generations to find the parent
            {
                parent = GetParent(parent, recursioncount+1);
            }
            return parent;
        }

        private void Panel_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.Move;
        }

        private void Panel_Drop(object sender, DragEventArgs e)
        {
            try
            {
                if (e.Handled == false)
                {
                    StackPanel dropped_on = (StackPanel)sender;
                    Border source_border = e.Data.GetData(typeof(Border)) as Border;
                    if (dropped_on != null && source_border != null)
                    {
                        StackPanel source_parent = GetParent(source_border) as StackPanel;
                        if (source_parent != null)
                        {
                            if (e.AllowedEffects.HasFlag(DragDropEffects.Move))
                            {
                                source_parent.Children.Remove(source_border);
                                dropped_on.Children.Add(source_border);
                                e.Effects = DragDropEffects.Move;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error dropping on panel: {ex.StackTrace}");
                MessageBox.Show($"Error dropping on panel: {ex.Message}.", "Panel Drop Error");
            }
            finally
            {
                e.Handled = true;
            }
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            CreateExcelReport();
        }

        private void CreateExcelReport()
        {
            #region Prep UI for Wait Screen
            ReportBorder.IsEnabled = false;
            AllBorder.IsEnabled = false;
            SelectedBorder.IsEnabled = false;
            Mouse.OverrideCursor = Cursors.Wait;
            LoadingRect.Visibility = Visibility.Visible;
            Loading.Visibility = Visibility.Visible;
            #endregion

            ExcelProcessor processor = new ExcelProcessor();
            processor.ProgressUpdated += Processor_ProgressUpdated;
            processor.ProcessingComplete += Processor_ProcessingComplete;
            processor.ExcelFromReport(ActiveReport);           
        }

        private void Processor_ProcessingComplete(object sender, EventArgs e)
        {
            MessageBox.Show($"Report saved to {ActiveReport.ExportFile}.", "Report Saved");
            System.Diagnostics.Process.Start(ActiveReport.ExportFile);

            #region Revert UI from Wait Screen
            SynchronizationContextObject.Post(
               new SendOrPostCallback(
                   o =>
                   {
                       ReportBorder.IsEnabled = true;
                       AllBorder.IsEnabled = true;
                       SelectedBorder.IsEnabled = true;
                       LoadingRect.Visibility = Visibility.Collapsed;
                       Loading.Visibility = Visibility.Collapsed;
                       Mouse.OverrideCursor = Cursors.Arrow;
                   }
                ),
               null
            );
            #endregion

        }

        private void Processor_ProgressUpdated(object sender, ProgressUpdatedArgs e)
        {
            UpdateUI(e.Message);
        }
        
        public void UpdateUI(string label)
        {
            SynchronizationContextObject.Post(
                new SendOrPostCallback(
                    o => { Loading.Content = label; }
                ), 
                label
            );
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Global.ReleaseExcelProcesses();
        }
    }
}
