using System;
using System.Collections.Generic;
using System.Web.Script.Serialization;

namespace ParameterTestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("=== 工具格式兼容性测试 ===\n");

            // 测试各种输入格式
            string[] testCases = new string[]
            {
                // 标准格式
                "{\"action\":[{\"type\":\"replace\",\"detail\":{\"original\":\"旧文本\",\"new\":\"新文本\"}}]}",

                // 兼容格式1：使用action字段替代type
                "{\"action\":[{\"action\":\"replace\",\"detail\":{\"original\":\"旧文本\",\"new\":\"新文本\"}}]}",

                // 兼容格式2：扁平参数结构
                "{\"action\":[{\"type\":\"replace\",\"original\":\"旧文本\",\"new\":\"新文本\"}]}",

                // 兼容格式3：action字段+扁平参数
                "{\"action\":[{\"action\":\"replace\",\"original\":\"旧文本\",\"new\":\"新文本\"}]}",

                // 兼容格式4：参数名映射
                "{\"action\":[{\"type\":\"replace\",\"detail\":{\"old_text\":\"旧文本\",\"new_text\":\"新文本\"}}]}"
            };

            string[] descriptions = new string[]
            {
                "标准格式 (type + detail)",
                "兼容格式1 (action字段替代type)",
                "兼容格式2 (扁平参数结构)",
                "兼容格式3 (action字段+扁平参数)",
                "兼容格式4 (参数名映射 old_text->original)"
            };

            for (int i = 0; i < testCases.Length; i++)
            {
                Console.WriteLine($"--- 测试用例 {i + 1}: {descriptions[i]} ---");
                Console.WriteLine($"输入: {testCases[i]}");
                TestFormat(testCases[i]);
                Console.WriteLine();
            }

            Console.WriteLine("按任意键退出...");
            Console.ReadKey();
        }

        static void TestFormat(string jsonInput)
        {
            try
            {
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                var root = serializer.Deserialize<Dictionary<string, object>>(jsonInput);

                if (root.ContainsKey("action"))
                {
                    var actions = root["action"] as System.Collections.ArrayList;
                    if (actions != null && actions.Count > 0)
                    {
                        var action = actions[0] as Dictionary<string, object>;
                        if (action != null)
                        {
                            // 应用参数映射逻辑（与McpTools.cs相同）
                            var normalizedAction = new Dictionary<string, object>();

                            // 确定操作类型
                            string operationType = null;
                            if (action.ContainsKey("type"))
                            {
                                operationType = action["type"]?.ToString();
                            }
                            else if (action.ContainsKey("action"))
                            {
                                operationType = action["action"]?.ToString();
                            }

                            if (string.IsNullOrEmpty(operationType))
                            {
                                Console.WriteLine("❌ 缺少操作类型");
                                return;
                            }

                            normalizedAction["type"] = operationType;

                            // 创建detail对象
                            var normalizedDetail = new Dictionary<string, object>();

                            // 处理detail字段（标准格式）
                            if (action.ContainsKey("detail"))
                            {
                                var detail = action["detail"] as Dictionary<string, object>;
                                if (detail != null)
                                {
                                    foreach (var kvp in detail)
                                    {
                                        if (kvp.Key == "old_text")
                                        {
                                            normalizedDetail["original"] = kvp.Value;
                                        }
                                        else if (kvp.Key == "new_text")
                                        {
                                            normalizedDetail["new"] = kvp.Value;
                                        }
                                        else
                                        {
                                            normalizedDetail[kvp.Key] = kvp.Value;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                // 处理扁平格式（所有参数直接在action对象下）
                                foreach (var kvp in action)
                                {
                                    if (kvp.Key == "type" || kvp.Key == "action")
                                        continue; // 跳过type字段

                                    if (kvp.Key == "old_text")
                                    {
                                        normalizedDetail["original"] = kvp.Value;
                                    }
                                    else if (kvp.Key == "new_text")
                                    {
                                        normalizedDetail["new"] = kvp.Value;
                                    }
                                    else
                                    {
                                        normalizedDetail[kvp.Key] = kvp.Value;
                                    }
                                }
                            }

                            normalizedAction["detail"] = normalizedDetail;

                            // 生成最终JSON
                            string finalJson = serializer.Serialize(normalizedAction);
                            Console.WriteLine($"输出: {finalJson}");

                            // 验证格式
                            var testRoot = serializer.Deserialize<Dictionary<string, object>>(finalJson);
                            string opType = testRoot["type"]?.ToString();

                            if (opType == "replace")
                            {
                                var testDetail = testRoot["detail"] as Dictionary<string, object>;
                                if (testDetail != null && testDetail.ContainsKey("original") && testDetail.ContainsKey("new"))
                                {
                                    Console.WriteLine("✅ 格式转换成功");
                                }
                                else
                                {
                                    Console.WriteLine("❌ 格式转换失败：缺少必需参数");
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("❌ action不是Dictionary对象");
                        }
                    }
                    else
                    {
                        Console.WriteLine("❌ action数组为空");
                    }
                }
                else
                {
                    Console.WriteLine("❌ 缺少action字段");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 解析失败: {ex.Message}");
            }
        }
    }
}