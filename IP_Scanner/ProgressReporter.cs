using System;
using System.Windows;
using System.Windows.Controls;

namespace IPProcessingTool
{
    public class ProgressReporter
    {
        private ProgressBar progressBar;
        private TextBlock statusText;
        private int totalItems;
        private int processedItems;
        private DateTime startTime;

        public ProgressReporter(ProgressBar progressBar, TextBlock statusText)
        {
            this.progressBar = progressBar;
            this.statusText = statusText;
        }

        public void Initialize(int totalItems)
        {
            this.totalItems = totalItems;
            processedItems = 0;
            startTime = DateTime.Now;
            UpdateUI();
        }

        public void IncrementProgress()
        {
            processedItems++;
            UpdateUI();
        }

        private void UpdateUI()
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                double progress = (double)processedItems / totalItems;
                progressBar.Value = progress * 100;

                TimeSpan elapsedTime = DateTime.Now - startTime;
                TimeSpan estimatedTotalTime = TimeSpan.FromTicks((long)(elapsedTime.Ticks / progress));
                TimeSpan remainingTime = estimatedTotalTime - elapsedTime;

                statusText.Text = $"Progress: {processedItems}/{totalItems} " +
                                  $"({progress:P0}) | " +
                                  $"Elapsed: {FormatTimeSpan(elapsedTime)} | " +
                                  $"Remaining: {FormatTimeSpan(remainingTime)}";
            });
        }

        private string FormatTimeSpan(TimeSpan timeSpan)
        {
            return timeSpan.ToString(@"hh\:mm\:ss");
        }
    }
}