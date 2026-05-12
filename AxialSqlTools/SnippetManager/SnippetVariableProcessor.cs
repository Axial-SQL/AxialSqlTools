using System;

namespace AxialSqlTools
{
    public class SnippetVariableResult
    {
        public string ProcessedText { get; set; }
        public int CursorOffset { get; set; }

        public SnippetVariableResult(string processedText, int cursorOffset)
        {
            ProcessedText = processedText;
            CursorOffset = cursorOffset;
        }
    }

    public static class SnippetVariableProcessor
    {
        public static SnippetVariableResult ProcessVariables(string body, string cursorMarker)
        {
            if (string.IsNullOrEmpty(body))
                return new SnippetVariableResult(string.Empty, -1);

            string text = body;

            // Replace built-in variables
            text = text.Replace("$DATE$", DateTime.Now.ToString("yyyy-MM-dd"));
            text = text.Replace("$TIME$", DateTime.Now.ToString("HH:mm:ss"));
            text = text.Replace("$USER$", Environment.UserName);

            // Find and remove cursor marker
            int cursorOffset = -1;
            if (!string.IsNullOrEmpty(cursorMarker))
            {
                int markerIndex = text.IndexOf(cursorMarker);
                if (markerIndex >= 0)
                {
                    cursorOffset = markerIndex;
                    text = text.Remove(markerIndex, cursorMarker.Length);
                }
            }

            return new SnippetVariableResult(text, cursorOffset);
        }
    }
}
