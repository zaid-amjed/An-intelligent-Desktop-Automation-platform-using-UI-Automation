using AutoDataEntryProject.Models;
using AutoDataEntryProject.Repositories;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace AutoDataEntryProject
{
    
    public partial class MainWindow : Window
    {
        private readonly AutomationManager _manager;
        private readonly IStudentRepository _studentRepository;
        private CancellationTokenSource? _cts;
        private List<Student> _students;
        private string _excelPath = string.Empty;
        private string _appPath = string.Empty;

        public MainWindow()
        {
            InitializeComponent();
            _manager = new AutomationManager();
            _studentRepository = new StudentExcelRepository(); 
            _students = new List<Student>();
            
            InitializeUI();
        }

        private void InitializeUI()
        {
            btnStart.IsEnabled = false;
            btnStop.IsEnabled = false;
            lblTotalCount.Text = "0";
            lblProcessedCount.Text = "0";
            lblElapsedTime.Text = "00:00";
            progressBar.Value = 0;
            lblProgress.Text = "جاهز للبدء";
            lblStatus.Text = "● جاهز";
            LogMessage("مرحباً بك في نظام إدخال البيانات التلقائي");
            LogMessage("يستخدم Repository Pattern للوصول للبيانات");
        }

       

        private void btnSelectExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Title = "اختر ملف Excel",
                    Filter = "Excel Files|*.xlsx;*.xlsm|All Files|*.*",
                    FilterIndex = 1
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    _excelPath = openFileDialog.FileName;
                    txtExcelPath.Text = _excelPath;
                    LogMessage($"تم اختيار ملف Excel: {Path.GetFileName(_excelPath)}");

                    LoadStudentsFromExcel();
                    CheckReadyToStart();
                }
            }
            catch (Exception ex)
            {
                ShowError($"خطأ في اختيار ملف Excel: {ex.Message}");
            }
        }

        private void btnSelectTargetApp_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Title = "اختر التطبيق الهدف",
                    Filter = "Executable Files|*.exe|All Files|*.*",
                    FilterIndex = 1
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    _appPath = openFileDialog.FileName;
                    txtAppPath.Text = _appPath;
                    LogMessage($"تم اختيار التطبيق الهدف: {Path.GetFileName(_appPath)}");
                    CheckReadyToStart();
                }
            }
            catch (Exception ex)
            {
                ShowError($"خطأ في اختيار التطبيق: {ex.Message}");
            }
        }

        

      
        private void LoadStudentsFromExcel()
        {
            try
            {
                LogMessage("جاري قراءة بيانات الطلاب باستخدام Repository Pattern...");
                
                // استخدام Repository للتحقق من صحة الملف
                if (!_studentRepository.ValidateDataSource(_excelPath))
                {
                    ShowError("ملف Excel غير صحيح. تأكد من وجود 3 أعمدة على الأقل");
                    _students.Clear();
                    lblTotalCount.Text = "0";
                    return;
                }

                // استخدام Repository لقراءة البيانات
                _students = _studentRepository.GetAllStudents(_excelPath);

                if (_students.Count == 0)
                {
                    ShowWarning("لم يتم العثور على بيانات في ملف Excel");
                    lblTotalCount.Text = "0";
                    return;
                }

                // التحقق من صحة البيانات
                int validCount = 0;
                int invalidCount = 0;
                foreach (var student in _students)
                {
                    if (student.IsValid())
                        validCount++;
                    else
                        invalidCount++;
                }

                lblTotalCount.Text = validCount.ToString();
                LogMessage($"✓ Repository Pattern: تم قراءة {_students.Count} سجل");
                LogMessage($"  - صحيح: {validCount} | غير صحيح: {invalidCount}");

                if (invalidCount > 0)
                {
                    ShowWarning($"تحذير: يوجد {invalidCount} سجل غير صحيح");
                }
            }
            catch (Exception ex)
            {
                ShowError($"خطأ في قراءة ملف Excel: {ex.Message}");
                _students.Clear();
                lblTotalCount.Text = "0";
            }
        }

       

        

        private async void btnStartAutomation_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_excelPath) || string.IsNullOrWhiteSpace(_appPath))
            {
                ShowWarning("الرجاء اختيار ملف Excel والتطبيق الهدف أولاً");
                return;
            }

            if (_students == null || _students.Count == 0)
            {
                ShowWarning("لا توجد بيانات للمعالجة");
                return;
            }

            _cts = new CancellationTokenSource();
            btnStart.IsEnabled = false;
            btnStop.IsEnabled = true;
            lblStatus.Text = "● جاري التشغيل...";
            lblProcessedCount.Text = "0";
            progressBar.Value = 0;

            LogMessage("════════════════════════════════════");
            LogMessage("بدء عملية الأتمتة...");
            LogMessage($"التطبيق الهدف: {Path.GetFileName(_appPath)}");
            LogMessage($"عدد السجلات: {_students.Count}");
            LogMessage("════════════════════════════════════");

            try
            {
                await _manager.RunAutomation(
                    _appPath,
                    _students,
                    UpdateProgress,
                    _cts.Token
                );

                LogMessage("════════════════════════════════════");
                LogMessage("✓ اكتملت العملية بنجاح!");
                LogMessage("════════════════════════════════════");
                lblStatus.Text = "● اكتمل";
                MessageBox.Show(
                    "تمت عملية إدخال البيانات بنجاح!",
                    "نجحت العملية",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information
                );
            }
            catch (OperationCanceledException)
            {
                LogMessage("════════════════════════════════════");
                LogMessage("✗ تم إلغاء العملية");
                LogMessage("════════════════════════════════════");
                lblStatus.Text = "● تم الإلغاء";
                MessageBox.Show(
                    "تم إلغاء العملية",
                    "إلغاء",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );
            }
            catch (Exception ex)
            {
                LogMessage("════════════════════════════════════");
                LogMessage($"✗ خطأ: {ex.Message}");
                LogMessage("════════════════════════════════════");
                lblStatus.Text = "● خطأ";
                ShowError($"حدث خطأ أثناء تنفيذ الأتمتة:\n\n{ex.Message}");
            }
            finally
            {
                btnStart.IsEnabled = true;
                btnStop.IsEnabled = false;
            }
        }

        private void btnStop_Click(object sender, RoutedEventArgs e)
        {
            if (_cts != null && !_cts.IsCancellationRequested)
            {
                var result = MessageBox.Show(
                    "هل أنت متأكد من إيقاف العملية؟",
                    "تأكيد الإيقاف",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question
                );

                if (result == MessageBoxResult.Yes)
                {
                    LogMessage("جاري إيقاف العملية...");
                    _cts.Cancel();
                    btnStop.IsEnabled = false;
                }
            }
        }

        private void UpdateProgress(AutomationProgress progress)
        {
            Dispatcher.Invoke(() =>
            {
                progressBar.Value = progress.ProgressPercentage;
                lblProgress.Text = $"التقدم: {progress.ProgressPercentage:F1}%";
                lblProcessedCount.Text = progress.CurrentCount.ToString();
                lblElapsedTime.Text = progress.ElapsedTime.ToString(@"mm\:ss");
                
                LogMessage(progress.GetFormattedStatus());

                if (progress.IsCompleted)
                {
                    lblStatus.Text = "● اكتمل";
                }
            });
        }

        

        

        private void CheckReadyToStart()
        {
            btnStart.IsEnabled = !string.IsNullOrWhiteSpace(_excelPath) &&
                                !string.IsNullOrWhiteSpace(_appPath) &&
                                _students != null &&
                                _students.Count > 0;
        }

        private void LogMessage(string message)
        {
            Dispatcher.Invoke(() =>
            {
                txtLog.AppendText($"{message}\n");
                txtLog.ScrollToEnd();
            });
        }

        private void ShowError(string message)
        {
            LogMessage($"✗ خطأ: {message}");
            MessageBox.Show(
                message,
                "خطأ",
                MessageBoxButton.OK,
                MessageBoxImage.Error
            );
        }

        private void ShowWarning(string message)
        {
            LogMessage($"⚠ تحذير: {message}");
            MessageBox.Show(
                message,
                "تحذير",
                MessageBoxButton.OK,
                MessageBoxImage.Warning
            );
        }

        

       

        protected override void OnClosing(CancelEventArgs e)
        {
            if (_cts != null && !_cts.IsCancellationRequested)
            {
                var result = MessageBox.Show(
                    "هناك عملية قيد التشغيل. هل تريد الخروج؟",
                    "تأكيد الخروج",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question
                );

                if (result == MessageBoxResult.No)
                {
                    e.Cancel = true;
                    return;
                }

                _cts.Cancel();
            }

            base.OnClosing(e);
        }

        
    }
}
