using System;

namespace SSMS_EnvTabs
{
    internal static class TabGroupColorSolver
    {
        // Solves for a salt that results in the target ColorIndex for the given base regex.
        // <param name="baseRegex">The base regex string (e.g. filename regex).</param>
        // <param name="targetColorIndex">The desired SSMS color index (0-15).</param>
        // <param name="forbiddenHashes">Optional set of regex hashes already recorded in the SSMS
        //   group-color JSON with a different color. Any candidate whose hash is in this set will
        //   be skipped — otherwise SSMS would apply its stale override and ignore our salt.</param>
        // <returns>A salt string (e.g. "123") that when appended as (?#salt:123) produces the target color.</returns>
        public static string Solve(string baseRegex, int targetColorIndex, System.Collections.Generic.ICollection<int> forbiddenHashes = null)
        {
            if (targetColorIndex < 0 || targetColorIndex > 15) return null;

            // Brute force numbers until we find a match.
            // Start from 1 up to reasonable limit.
            // Usually finds a match within a few dozen attempts (probability 1/16).
            for (int i = 1; i < 10000; i++)
            {
                string salt = i.ToString();
                // Construct the exact string SSMS sees in the config file
                string fullRegex = baseRegex + "(?#salt:" + salt + ")";
                
                int hash = GetSsmsStableHashCode(fullRegex);
                int calculatedIndex = Math.Abs(hash) % 16;
                
                if (calculatedIndex == targetColorIndex)
                {
                    // Skip this candidate if SSMS already has this hash mapped to a different
                    // color in its override JSON — SSMS would just re-apply that stale color.
                    if (forbiddenHashes != null && forbiddenHashes.Contains(hash))
                    {
                        continue;
                    }

                    return salt;
                }
            }
            
            return null; // Should essentially never happen
        }

        public static string SolveForColor(string baseRegex, int targetColorIndex)
        {
            string salt = Solve(baseRegex, targetColorIndex);
            if (salt != null)
            {
                return baseRegex + "(?#salt:" + salt + ")";
            }
            return baseRegex;
        }

        // Calculates the hash code using the Legacy .NET 32-bit x86 string hashing algorithm.
        // This matches what SSMS versions (based on older .NET Framework execution) use.
        public static int GetSsmsStableHashCode(string str)
        {
            unchecked
            {
                int hash1 = 5381;
                int hash2 = hash1;

                for (int i = 0; i < str.Length; i += 2)
                {
                    hash1 = ((hash1 << 5) + hash1) ^ str[i];
                    if (i == str.Length - 1) break;
                    hash2 = ((hash2 << 5) + hash2) ^ str[i + 1];
                }

                return hash1 + (hash2 * 1566083941);
            }
        }

    }
}
