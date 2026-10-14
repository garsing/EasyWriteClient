using System;

namespace WordDocumentRoundtripTest
{
    internal sealed class TestRun
    {
        public int CasesPassed { get; private set; }
        public int CasesFailed { get; private set; }
        public int CasesSkipped { get; private set; }

        public int TotalTarget { get; set; }

        private string Progress()
        {
            return TotalTarget > 0
                ? (CasesPassed + CasesFailed + CasesSkipped) + "/" + TotalTarget
                : "?";
        }

        public void NotePassed(int count)
        {
            if (count > 0)
            {
                CasesPassed += count;
            }
        }

        public void CaseOk(string name)
        {
            CasesPassed++;
            Console.WriteLine("  [OK]   " + Progress().PadRight(7) + " " + name);
        }

        public void CaseFail(string name, string reason)
        {
            CasesFailed++;
            Console.WriteLine("  [FAIL] " + Progress().PadRight(7) + " " + name);
            Console.WriteLine("         " + reason);
        }

        public void CaseSkip(string name, string reason)
        {
            CasesSkipped++;
            Console.WriteLine("  [SKIP] " + Progress().PadRight(7) + " " + name + " — " + reason);
        }
    }
}
