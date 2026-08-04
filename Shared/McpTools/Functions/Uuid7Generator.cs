using System;

namespace WordAddIn1
{
    /// <summary>
    /// UUID7 生成工具类
    /// UUID7 是顺序型 UUID，基于时间戳生成，适合作为数据库主键
    /// </summary>
    public static class Uuid7Generator
    {
        /// <summary>
        /// 生成 UUID7（顺序型 UUID）
        /// UUID7 格式：48位时间戳（毫秒） + 4位版本(7) + 12位变体 + 62位随机数
        /// 参考：RFC 4122 UUID Version 7 (draft)
        /// 格式：tttttttt-tttt-7vvv-rrrr-rrrrrrrrrrrr
        /// </summary>
        /// <returns>UUID7 字符串</returns>
        public static string GenerateUuid7()
        {
            // 获取当前时间戳（毫秒）
            long timestampMs = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
            
            // 48位时间戳，转换为12个16进制字符
            string timestampHex = (timestampMs & 0xFFFFFFFFFFFFL).ToString("x12").PadLeft(12, '0');
            
            // 版本号 7 (4位)
            string versionHex = "7";
            
            // 变体（12位），最高位为1（表示RFC 4122变体）
            Random random = new Random();
            int variantBits = 0x8 | (random.Next(0, 2048)); // 最高位为1，其余11位随机
            string variantHex = variantBits.ToString("x3").PadLeft(3, '0');
            
            // 62位随机数
            long randBits = ((long)random.Next(0, int.MaxValue) << 31) | (long)random.Next(0, int.MaxValue);
            randBits &= 0x3FFFFFFFFFFFFFFF; // 确保只有62位
            string randHex = randBits.ToString("x15").PadLeft(15, '0');
            
            // 组合成 UUID7 格式：xxxxxxxx-xxxx-7vvv-rrrr-rrrrrrrrrrrr
            string part1 = timestampHex.Substring(0, 8);  // 前8个字符（32位时间戳）
            string part2 = timestampHex.Substring(8, 4);  // 后4个字符（16位时间戳）
            string part3 = versionHex + variantHex.Substring(0, 2);  // 版本号 + 变体前2位
            string part4 = variantHex.Substring(2, 1) + randHex.Substring(0, 3);  // 变体第3位 + 随机数前3位
            string part5 = randHex.Substring(3);  // 随机数剩余部分
            
            return $"{part1}-{part2}-{part3}-{part4}-{part5}";
        }
    }
}

