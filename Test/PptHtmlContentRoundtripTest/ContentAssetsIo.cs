using System;
using System.IO;

namespace PptHtmlContentRoundtripTest
{
    internal static class ContentAssetsIo
    {
        public static readonly byte[] MinimalPngRed = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

        public static readonly byte[] MinimalPngBlue = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPj/HwADBwIAMCbHYQAAAABJRU5ErkJggg==");

        public static readonly byte[] MinimalPngGreen = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");

        public static ContentAssets Ensure(string dir)
        {
            Directory.CreateDirectory(dir);
            string red = Path.Combine(dir, "red.png");
            string blue = Path.Combine(dir, "blue.png");
            string green = Path.Combine(dir, "green.png");
            File.WriteAllBytes(red, MinimalPngRed);
            File.WriteAllBytes(blue, MinimalPngBlue);
            File.WriteAllBytes(green, MinimalPngGreen);
            return new ContentAssets
            {
                RedPng = Path.GetFullPath(red),
                BluePng = Path.GetFullPath(blue),
                GreenPng = Path.GetFullPath(green)
            };
        }
    }
}
