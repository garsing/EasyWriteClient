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

        public static ContentAssets Ensure(string dir)
        {
            Directory.CreateDirectory(dir);
            string red = Path.Combine(dir, "red.png");
            string blue = Path.Combine(dir, "blue.png");
            File.WriteAllBytes(red, MinimalPngRed);
            File.WriteAllBytes(blue, MinimalPngBlue);
            return new ContentAssets
            {
                RedPng = Path.GetFullPath(red),
                BluePng = Path.GetFullPath(blue)
            };
        }
    }
}
