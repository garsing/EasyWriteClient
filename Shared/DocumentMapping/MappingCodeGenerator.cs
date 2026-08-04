namespace WordAddIn1.DocumentMapping
{
    internal static class MappingCodeGenerator
    {
        public static string GenerateSentenceName(int index) => GenerateCode("S_", index);
        public static string GenerateParagraphName(int index) => GenerateCode("P_", index);

        private static string GenerateCode(string prefix, int index)
        {
            string name = "";
            int idx = index;
            for (int i = 0; i < 5; i++)
            {
                int charCode = idx % 62;
                char ch;
                if (charCode < 10)
                {
                    ch = (char)('0' + charCode);
                }
                else if (charCode < 36)
                {
                    ch = (char)('a' + (charCode - 10));
                }
                else
                {
                    ch = (char)('A' + (charCode - 36));
                }

                name = ch + name;
                idx /= 62;
            }

            return prefix + name;
        }
    }
}
