namespace StreamingNotificationsSample
{
    public static class Utils
    {
        public static void ShowTaskbarNotifier(string msg)
        {
            var taskBarNotifier = new TaskbarNotifier();
            taskBarNotifier.SetBackgroundBitmap();
            taskBarNotifier.SetCloseBitmap();
            taskBarNotifier.KeepVisibleOnMouseOver = true;
            taskBarNotifier.ReShowOnMouseOver = true;
            taskBarNotifier.Show("ConnectWise\nMessage", msg, 500, 3000, 500);
        }
    }
}
