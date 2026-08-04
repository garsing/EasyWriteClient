using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace LlmClientTest
{
    class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("=== LlmClient 测试程序 ===\n");

            // 测试1: 简单对话测试
            Console.WriteLine("1. 测试简单对话:");
            try
            {
                string result1 = await WordAddIn1.LlmExample.TestChatAsync();
                Console.WriteLine("✓ 成功: " + result1);
            }
            catch (Exception ex)
            {
                Console.WriteLine("✗ 失败: " + ex.Message);
            }

            Console.WriteLine();

            // 测试2: 多轮对话测试
            Console.WriteLine("2. 测试多轮对话:");
            try
            {
                string result2 = await WordAddIn1.LlmExample.TestMultiTurnChatAsync();
                Console.WriteLine("✓ 成功: " + result2);
            }
            catch (Exception ex)
            {
                Console.WriteLine("✗ 失败: " + ex.Message);
            }

            Console.WriteLine();

            // 测试3: 自定义测试
            Console.WriteLine("3. 自定义测试:");
            try
            {
                var client = new WordAddIn1.LlmClient("sk-2e2929af6bde429990eef75c39e6afd7");

                // 测试不同的提示
                var messages = new List<WordAddIn1.ChatMessage>
                {
                    new WordAddIn1.ChatMessage { role = "system", content = "你是一个编程助手，请用简洁的语言回答。" },
                    new WordAddIn1.ChatMessage { role = "user", content = "用C#写一个Hello World程序" }
                };

                var response = await client.ChatAsync(messages);
                Console.WriteLine("✓ 自定义测试成功:");
                Console.WriteLine(response.choices[0].message.content);

                // 显示Token使用情况
                Console.WriteLine($"\nToken使用情况:");
                Console.WriteLine($"  输入: {response.usage.prompt_tokens} tokens");
                Console.WriteLine($"  输出: {response.usage.completion_tokens} tokens");
                Console.WriteLine($"  总计: {response.usage.total_tokens} tokens");
            }
            catch (WordAddIn1.LlmException ex)
            {
                Console.WriteLine("✗ API调用失败: " + ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("✗ 未知错误: " + ex.Message);
            }

            Console.WriteLine("\n=== 测试完成 ===");
            Console.WriteLine("按任意键退出...");
            Console.ReadKey();
        }
    }
}
