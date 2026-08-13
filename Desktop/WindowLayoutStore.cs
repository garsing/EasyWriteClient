using System;
using System.Drawing;
using System.IO;
using Newtonsoft.Json;

namespace EasyWriteClient.Desktop
{
    internal enum WindowLayoutMode
    {
        Expanded = 0,
        Compact = 1
    }

    /// <summary>
    /// 持久化 Desktop 窗口形态：%LOCALAPPDATA%\EasyWriteDesktop\window-layout.json
    /// </summary>
    internal static class WindowLayoutStore
    {
        private sealed class LayoutDto
        {
            [JsonProperty("mode")]
            public string Mode { get; set; }

            [JsonProperty("compactX")]
            public int? CompactX { get; set; }

            [JsonProperty("compactY")]
            public int? CompactY { get; set; }

            [JsonProperty("compactW")]
            public int? CompactW { get; set; }

            [JsonProperty("compactH")]
            public int? CompactH { get; set; }

            /// <summary>缩小版闲置自动收为悬浮球；缺省 true。本期无设置页 UI。</summary>
            [JsonProperty("autoFloatEnabled")]
            public bool? AutoFloatEnabled { get; set; }

            [JsonProperty("ballX")]
            public int? BallX { get; set; }

            [JsonProperty("ballY")]
            public int? BallY { get; set; }
        }

        public static string FilePath
        {
            get
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "EasyWriteDesktop");
                return Path.Combine(dir, "window-layout.json");
            }
        }

        public static WindowLayoutMode LoadMode(WindowLayoutMode fallback = WindowLayoutMode.Expanded)
        {
            LayoutDto dto = ReadDto();
            if (dto == null)
            {
                return fallback;
            }

            if (string.Equals(dto.Mode, "compact", StringComparison.OrdinalIgnoreCase))
            {
                return WindowLayoutMode.Compact;
            }

            if (string.Equals(dto.Mode, "expanded", StringComparison.OrdinalIgnoreCase))
            {
                return WindowLayoutMode.Expanded;
            }

            return fallback;
        }

        /// <summary>若曾保存过缩小版 Bounds 则返回，否则 null。</summary>
        public static Rectangle? LoadCompactBounds()
        {
            LayoutDto dto = ReadDto();
            if (dto?.CompactW == null || dto.CompactH == null
                || dto.CompactX == null || dto.CompactY == null)
            {
                return null;
            }

            if (dto.CompactW.Value < 200 || dto.CompactH.Value < 200)
            {
                return null;
            }

            return new Rectangle(
                dto.CompactX.Value,
                dto.CompactY.Value,
                dto.CompactW.Value,
                dto.CompactH.Value);
        }

        public static bool GetAutoFloatEnabled()
        {
            LayoutDto dto = ReadDto();
            if (dto?.AutoFloatEnabled == null)
            {
                return true;
            }

            return dto.AutoFloatEnabled.Value;
        }

        public static void SetAutoFloatEnabled(bool enabled)
        {
            try
            {
                LayoutDto dto = ReadDto() ?? new LayoutDto();
                dto.AutoFloatEnabled = enabled;
                WriteDto(dto);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[WindowLayoutStore] SetAutoFloatEnabled: " + ex.Message);
            }
        }

        public static Point? LoadBallLocation()
        {
            LayoutDto dto = ReadDto();
            if (dto?.BallX == null || dto.BallY == null)
            {
                return null;
            }

            return new Point(dto.BallX.Value, dto.BallY.Value);
        }

        public static void SaveBallLocation(Point location)
        {
            try
            {
                LayoutDto dto = ReadDto() ?? new LayoutDto();
                dto.BallX = location.X;
                dto.BallY = location.Y;
                WriteDto(dto);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[WindowLayoutStore] SaveBallLocation: " + ex.Message);
            }
        }

        public static void Save(WindowLayoutMode mode, Rectangle? compactBounds = null)
        {
            try
            {
                LayoutDto dto = ReadDto() ?? new LayoutDto();
                dto.Mode = mode == WindowLayoutMode.Compact ? "compact" : "expanded";
                if (compactBounds.HasValue)
                {
                    Rectangle r = compactBounds.Value;
                    dto.CompactX = r.X;
                    dto.CompactY = r.Y;
                    dto.CompactW = r.Width;
                    dto.CompactH = r.Height;
                }

                WriteDto(dto);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[WindowLayoutStore] Save: " + ex.Message);
            }
        }

        /// <summary>兼容旧调用：只更新 mode。</summary>
        public static void Save(WindowLayoutMode mode)
        {
            Save(mode, null);
        }

        private static void WriteDto(LayoutDto dto)
        {
            string path = FilePath;
            string dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            File.WriteAllText(path, JsonConvert.SerializeObject(dto, Formatting.Indented));
        }

        private static LayoutDto ReadDto()
        {
            try
            {
                string path = FilePath;
                if (!File.Exists(path))
                {
                    return null;
                }

                string json = File.ReadAllText(path);
                return JsonConvert.DeserializeObject<LayoutDto>(json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[WindowLayoutStore] Read: " + ex.Message);
                return null;
            }
        }
    }
}
