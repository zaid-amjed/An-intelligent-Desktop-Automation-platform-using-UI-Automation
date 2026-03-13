using AutoDataEntryProject.Models;
using AutoDataEntryProject.Utilities;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Logging;
using FlaUI.UIA3;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace AutoDataEntryProject
{
    public class AutomationManager
    {
        #region Win32 API Imports (لإدارة النوافذ والحماية)
        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        private static extern bool SetWindowDisplayAffinity(IntPtr hWnd, uint dwAffinity);

        private static readonly IntPtr HWND_BOTTOM = new IntPtr(1);
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint WDA_EXCLUDEFROMCAPTURE = 0x00000011;
        #endregion

        private const int MAX_RETRY_ATTEMPTS = 3;
        private const int RETRY_DELAY_MS = 500;
        private const int TYPING_DELAY_MS = 30;

        public AutomationManager()
        {
            FlaUI.Core.Logging.Logger.Default.SetLevel(LogLevel.Error);
        }

        private async Task TypeTextSlowly(FlaUI.Core.AutomationElements.TextBox textBox, string text, CancellationToken token)
        {
            textBox.Text = "";
            await Task.Delay(50, token);
            
            foreach (char c in text)
            {
                if (token.IsCancellationRequested)
                    return;
                    
                textBox.Text += c;
                await Task.Delay(TYPING_DELAY_MS, token);
            }
        }

        public async Task RunAutomation(
            string targetAppPath,
            List<Student> students,
            Action<AutomationProgress> progressCallback,
            CancellationToken token)
        {
            if (string.IsNullOrWhiteSpace(targetAppPath))
                throw new ArgumentException("مسار التطبيق الهدف غير صحيح", nameof(targetAppPath));

            if (students == null || students.Count == 0)
                throw new ArgumentException("قائمة الطلاب فارغة", nameof(students));

            var perf = new PerformanceMonitor();
            perf.Start();

            Process? appProcess = null;
            Application? app = null;

            try
            {
                progressCallback?.Invoke(new AutomationProgress
                {
                    CurrentCount = 0,
                    TotalCount = students.Count,
                    Status = "جاري تشغيل التطبيق الهدف...",
                    ElapsedTime = perf.GetElapsed()
                });

                app = Application.Launch(targetAppPath);

                using (var automation = new UIA3Automation())
                {
                    var desktop = automation.GetDesktop();
                    Window? window = null;

                    progressCallback?.Invoke(new AutomationProgress
                    {
                        CurrentCount = 0,
                        TotalCount = students.Count,
                        Status = "البحث عن نافذة التطبيق...",
                        ElapsedTime = perf.GetElapsed()
                    });

                    for (int i = 0; i < 20 && window == null; i++)
                    {
                        if (token.IsCancellationRequested)
                        {
                            CleanupApplication(app, appProcess);
                            return;
                        }

                        window = desktop.FindFirstChild(cf => cf.ByName("نظام إدارة الطلاب"))?.AsWindow();
                        if (window == null)
                            await Task.Delay(500, token);
                    }

                    if (window == null)
                    {
                        CleanupApplication(app, appProcess);
                        throw new Exception("لم يتم العثور على نافذة التطبيق الهدف. تأكد من أن اسم النافذة هو 'نظام إدارة الطلاب'");
                    }

                    IntPtr handle = window.Properties.NativeWindowHandle.ValueOrDefault;
                    if (handle != IntPtr.Zero)
                    {
                        try
                        {
                            SetWindowDisplayAffinity(handle, WDA_EXCLUDEFROMCAPTURE);
                            SetWindowPos(handle, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"تحذير: فشل تطبيق إعدادات النافذة: {ex.Message}");
                        }
                    }

                    int successCount = 0;
                    int failedCount = 0;
                    List<string> errors = new List<string>();

                    for (int i = 0; i < students.Count; i++)
                    {
                        if (token.IsCancellationRequested)
                        {
                            progressCallback?.Invoke(new AutomationProgress
                            {
                                CurrentCount = successCount,
                                TotalCount = students.Count,
                                Status = "تم إلغاء العملية من قبل المستخدم",
                                ElapsedTime = perf.GetElapsed()
                            });
                            break;
                        }

                        var student = students[i];

                        if (!student.IsValid())
                        {
                            failedCount++;
                            errors.Add($"بيانات غير صحيحة للصف {i + 1}: {student}");
                            continue;
                        }

                        bool success = await ProcessStudentWithRetry(window, student, token);

                        if (success)
                        {
                            successCount++;
                            progressCallback?.Invoke(new AutomationProgress
                            {
                                CurrentCount = successCount,
                                TotalCount = students.Count,
                                Status = $"✓ تم إدخال: {student.Name}",
                                ElapsedTime = perf.GetElapsed()
                            });
                        }
                        else
                        {
                            failedCount++;
                            errors.Add($"فشل إدخال: {student.Name}");
                            progressCallback?.Invoke(new AutomationProgress
                            {
                                CurrentCount = successCount,
                                TotalCount = students.Count,
                                Status = $"✗ فشل إدخال: {student.Name}",
                                ElapsedTime = perf.GetElapsed()
                            });
                        }

                        await Task.Delay(300, token);
                    }

                    string finalStatus = $"اكتمل! النجاح: {successCount} | الفشل: {failedCount}";
                    if (errors.Count > 0)
                        finalStatus += $"\nأخطاء: {string.Join(", ", errors)}";

                    progressCallback?.Invoke(new AutomationProgress
                    {
                        CurrentCount = successCount,
                        TotalCount = students.Count,
                        Status = finalStatus,
                        ElapsedTime = perf.GetElapsed()
                    });
                }
            }
            catch (OperationCanceledException)
            {
                progressCallback?.Invoke(new AutomationProgress
                {
                    Status = "تم إلغاء العملية",
                    ElapsedTime = perf.GetElapsed()
                });
                throw;
            }
            catch (Exception ex)
            {
                progressCallback?.Invoke(new AutomationProgress
                {
                    Status = $"خطأ: {ex.Message}",
                    ElapsedTime = perf.GetElapsed()
                });
                throw;
            }
            finally
            {
                perf.Stop();
                CleanupApplication(app, appProcess);
            }
        }

        private async Task<bool> ProcessStudentWithRetry(Window window, Student student, CancellationToken token)
        {
            for (int attempt = 0; attempt < MAX_RETRY_ATTEMPTS; attempt++)
            {
                if (token.IsCancellationRequested)
                    return false;

                try
                {
                    var nameField = window.FindFirstDescendant(cf => cf.ByAutomationId("txtNameId"))?.AsTextBox();
                    var idField = window.FindFirstDescendant(cf => cf.ByAutomationId("txtIdId"))?.AsTextBox();
                    var emailField = window.FindFirstDescendant(cf => cf.ByAutomationId("txtEmailId"))?.AsTextBox();
                    var saveButton = window.FindFirstDescendant(cf => cf.ByAutomationId("btnSaveId"))?.AsButton();

                    if (nameField == null || idField == null || emailField == null || saveButton == null)
                    {
                        if (attempt < MAX_RETRY_ATTEMPTS - 1)
                        {
                            await Task.Delay(RETRY_DELAY_MS, token);
                            continue;
                        }
                        return false;
                    }

                    await TypeTextSlowly(nameField, student.Name, token);
                    await TypeTextSlowly(idField, student.StudentId, token);
                    await TypeTextSlowly(emailField, student.Email, token);

                    await Task.Delay(150, token);

                    if (saveButton.Patterns.Invoke.IsSupported)
                    {
                        saveButton.Invoke();
                    }
                    else if (saveButton.Patterns.LegacyIAccessible.IsSupported)
                    {
                        saveButton.Patterns.LegacyIAccessible.Pattern.DoDefaultAction();
                    }
                    else
                    {
                        saveButton.Click();
                    }

                    return true;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"محاولة {attempt + 1} فشلت: {ex.Message}");
                    if (attempt < MAX_RETRY_ATTEMPTS - 1)
                        await Task.Delay(RETRY_DELAY_MS, token);
                }
            }

            return false;
        }

        private void CleanupApplication(Application? app, Process? process)
        {
            try
            {
                if (app != null)
                {
                    app.Close();
                    app.Dispose();
                }
                else if (process != null && !process.HasExited)
                {
                    process.Kill();
                    process.Dispose();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"تحذير أثناء تنظيف التطبيق: {ex.Message}");
            }
        }
    }
}
