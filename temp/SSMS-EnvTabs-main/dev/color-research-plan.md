# SSMS Tab Color Research Plan

## Objective
Reverse engineer the hashing algorithm SSMS uses to assign colors to tab groups based on their regex patterns. This will allow us to programmatically generate "salted" regexes that target specific colors.

## The Theory
SSMS likely uses a deterministic hash of the regex string (e.g., `string.GetHashCode()`, CRC32, MD5) to generate a numeric `GroupId`. That `GroupId` is then mapped to a `ColorIndex` (1-16) likely via a modulo operation (e.g., `GroupId % 16`).

If we can verify the hash function and the mapping function, we can implement a "Solve for Color" feature:
1. Take the base regex.
2. Iterate through salt numbers (1, 2, 3...).
3. Hash `regex + salt`.
4. Check if result maps to target color.
5. Stop when match found.

## Data Collection Plan
To prove the functions, we need a dataset of ~10-20 triplets:
`{ RegexString, GroupId, ColorIndex }`

### Proposed Workflow (Manual Salt)

**Prerequisites:**
1. Launch SSMS (with extension installed).
2. Open `C:\Users\blake\Downloads\myQuery1.sql` (or similar).
3. Open Regex File: Click settings wheel (MIDS) -> "Configure regular expressions". This opens `ColorByRegexConfig.txt` in a tab.

**Data Collection Loop:**
1.  **Manual Salt**: Edit the `myQuery1.sql` line in `ColorByRegexConfig.txt` to add a random regex comment salt:
    `...myQuery1\.sql)$(?#salt:random_123)`
2.  **Save**: Save the file (Ctrl+S).
3.  **Trigger Write**: Right-click the **`myQuery1.sql` tab itself** (not the config tab!) -> "Set Tab Color" -> Pick the **SAME** color it currently has.
    *   *Observation*: This creates a NEW group entry in the SSMS internal JSON file.
4.  **Capture**: Press **"Capture Data"** button (Tools > SSMS EnvTabs > Capture Data).
    *   Extension detects the *new* group ID since last capture.
    *   Prompts user to enter the salt string used.
    *   Logs: `{ ReconstructedRegex, GroupId, ColorIndex }`.
5.  Repeat.

## Next Steps
Update "Capture Data" logic to simpler single-group assumption and prompt for salt.
