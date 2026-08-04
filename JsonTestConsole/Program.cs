using System;
using System.Collections.Generic;
using System.Web.Script.Serialization;

namespace JsonTestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("=== JSON处理测试 ===\n");

            // 测试JSON字符串清理
            string testJson = "{\"action\":[{\"type\":\"replace\",\"detail\":{\"old_text\":\"人工智能的核心技术包括机器学习、深度学习、自然语言处理和计算机视觉等。通过大量数据和算法训练，AI能够识别图像、理解语言、预测趋势和生成内容。例如，语音助手、自动驾驶和推荐系统都是AI的应用。\",\"new_text\":\"人工智能的核心技术涵盖机器学习、深度学习、自然语言处理以及计算机视觉等多个领域。通过海量数据和先进算法的持续训练，AI系统能够实现图像识别、语言理解、趋势预测乃至内容创作等复杂功能。例如，我们日常生活中广泛使用的语音助手、自动驾驶技术以及个性化推荐系统，都是人工智能技术的典型应用场景。\"}}]}文档修改已完成！我已经将润色后的内容直接替换了原文。";

            Console.WriteLine("原始JSON字符串:");
            Console.WriteLine(testJson);
            Console.WriteLine();

            string cleanedJson = CleanJsonString(testJson);
            Console.WriteLine("清理后的JSON字符串:");
            Console.WriteLine(cleanedJson);
            Console.WriteLine();

            // 解析清理后的JSON
            try
            {
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                var root = serializer.Deserialize<Dictionary<string, object>>(cleanedJson);

                if (root.ContainsKey("action"))
                {
                    var actions = root["action"] as System.Collections.ArrayList;
                    if (actions != null && actions.Count > 0)
                    {
                        var action = actions[0] as Dictionary<string, object>;
                        if (action != null)
                        {
                            Console.WriteLine("解析的action对象:");
                            foreach (var kvp in action)
                            {
                                Console.WriteLine($"  {kvp.Key}: {kvp.Value}");
                            }
                            Console.WriteLine();

                            // 规范化参数名
                            var normalizedAction = NormalizeAction(action);
                            Console.WriteLine("规范化后的action对象:");
                            foreach (var kvp in normalizedAction)
                            {
                                Console.WriteLine($"  {kvp.Key}: {kvp.Value}");
                            }
                            Console.WriteLine();

                            // 序列化回JSON
                            string finalJson = serializer.Serialize(normalizedAction);
                            Console.WriteLine("最终JSON字符串:");
                            Console.WriteLine(finalJson);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"解析失败: {ex.Message}");
            }

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 清理JSON字符串，确保它是有效的JSON格式
        /// </summary>
        static string CleanJsonString(string jsonString)
        {
            if (string.IsNullOrEmpty(jsonString))
                return jsonString;

            // 查找第一个完整的JSON对象（从{开始到}结束）
            int startIndex = jsonString.IndexOf('{');
            if (startIndex == -1)
                return jsonString;

            int braceCount = 0;
            int endIndex = -1;

            for (int i = startIndex; i < jsonString.Length; i++)
            {
                if (jsonString[i] == '{')
                {
                    braceCount++;
                }
                else if (jsonString[i] == '}')
                {
                    braceCount--;
                    if (braceCount == 0)
                    {
                        endIndex = i;
                        break;
                    }
                }
            }

            if (endIndex != -1 && endIndex >= startIndex)
            {
                string cleanJson = jsonString.Substring(startIndex, endIndex - startIndex + 1);
                Console.WriteLine($"清理后的JSON: {cleanJson}");
                return cleanJson;
            }

            return jsonString;
        }

        /// <summary>
        /// 处理参数名映射
        /// </summary>
        static Dictionary<string, object> NormalizeAction(Dictionary<string, object> action)
        {
            var normalizedAction = new Dictionary<string, object>();
            foreach (var kvp in action)
            {
                normalizedAction[kvp.Key] = kvp.Value;
            }

            // 如果是replace操作，处理参数名映射
            if (action.ContainsKey("type") && action["type"].ToString() == "replace")
            {
                if (action.ContainsKey("detail"))
                {
                    var detail = action["detail"] as Dictionary<string, object>;
                    if (detail != null)
                    {
                        var normalizedDetail = new Dictionary<string, object>();
                        foreach (var kvp in detail)
                        {
                            // 将old_text映射为original
                            if (kvp.Key == "old_text")
                            {
                                normalizedDetail["original"] = kvp.Value;
                            }
                            else
                            {
                                normalizedDetail[kvp.Key] = kvp.Value;
                            }
                        }
                        normalizedAction["detail"] = normalizedDetail;
                    }
                }
            }

            return normalizedAction;
        }
    }
}
