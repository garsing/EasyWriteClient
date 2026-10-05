using System;

namespace WordDocumentRoundtripTest
{
    internal sealed class TestRun
    {
        public int CasesPassed { get; private set; }
        public int CasesFailed { get; private set; }
        public int CasesSkipped { get; private set; }

        public void CaseOk(string name)
        {
            CasesPassed++;
            Console.WriteLine("  [OK] " + name);
        }

        public void CaseFail(string name, string reason)
        {
            CasesFailed++;
            Console.WriteLine("  [FAIL] " + name);
            Console.WriteLine("         " + reason);
        }

        public void CaseSkip(string name, string reason)
        {
            CasesSkipped++;
            Console.WriteLine("  [SKIP] " + name + " — " + reason);
        }
    }
}
