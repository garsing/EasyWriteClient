using System;

namespace PptChartRoundtripTest
{
    internal sealed class TestRun
    {
        public int Passed { get; private set; }

        public int Failed { get; private set; }

        public int Skipped { get; private set; }

        public void Ok(string name)
        {
            Passed++;
            Console.WriteLine("  [OK] " + name);
        }

        public void Fail(string name, string reason)
        {
            Failed++;
            Console.WriteLine("  [FAIL] " + name);
            Console.WriteLine("         " + reason);
        }

        public void Skip(string name, string reason)
        {
            Skipped++;
            Console.WriteLine("  [SKIP] " + name + " — " + reason);
        }

        public void Expect(string name, bool condition, string reason)
        {
            if (condition)
            {
                Ok(name);
            }
            else
            {
                Fail(name, reason);
            }
        }

        public void ExpectEqual(string name, string expected, string actual)
        {
            if (string.Equals(expected ?? "", actual ?? "", StringComparison.Ordinal))
            {
                Ok(name);
            }
            else
            {
                Fail(name, "期望 " + (expected ?? "(null)") + " ，实际 " + (actual ?? "(null)"));
            }
        }

        public void ExpectClose(string name, double expected, string actualRaw, double eps = 1e-6)
        {
            if (double.TryParse(actualRaw, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out double actual)
                && Math.Abs(expected - actual) <= eps)
            {
                Ok(name);
            }
            else
            {
                Fail(name, "期望 " + expected.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    + " ，实际 " + (actualRaw ?? "(null)"));
            }
        }
    }
}
