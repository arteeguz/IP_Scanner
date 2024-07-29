using System;
using System.Windows;
using System.Management;

namespace IPProcessingTool
{
    public static class ErrorHandler
    {
        public static void HandleException(Exception ex, string context, string userFriendlyMessage = null)
        {
            // Log the error
            Logger.Log(LogLevel.ERROR, ex.Message, context: context, additionalInfo: ex.StackTrace);

            // Display a user-friendly message
            string message = userFriendlyMessage ?? "An unexpected error occurred. Please try again or contact support if the problem persists.";
            MessageBox.Show(message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        public static void HandleWMIException(ManagementException ex, string ip)
        {
            string userFriendlyMessage = $"Unable to retrieve information from {ip}. This may be due to insufficient permissions or WMI configuration issues.";
            HandleException(ex, "WMI Error", userFriendlyMessage);
        }

        public static void HandleNetworkException(Exception ex, string ip)
        {
            string userFriendlyMessage = $"Unable to connect to {ip}. Please check the network connection and ensure the target machine is reachable.";
            HandleException(ex, "Network Error", userFriendlyMessage);
        }

        public static void HandleFileAccessException(Exception ex, string filePath)
        {
            string userFriendlyMessage = $"Unable to access the file: {filePath}. Please ensure you have the necessary permissions and the file is not in use by another program.";
            HandleException(ex, "File Access Error", userFriendlyMessage);
        }
    }
}