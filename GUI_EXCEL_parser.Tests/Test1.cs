using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;

namespace GUI_EXCEL_parser.Tests;

[TestClass]
public sealed class MechanismModelResolverTests
{
    [TestMethod]
    public void GetObjectModel_ReturnsDefaultForEmpty()
    {
        var model = MechanismModelResolver.GetObjectModel(string.Empty);
        Assert.AreEqual("AKKUYU_NO_MODEL", model);
    }

    [TestMethod]
    public void GetObjectModel_ReturnsValveForAaSgk()
    {
        var model = MechanismModelResolver.GetObjectModel("10SGK33AA007");
        Assert.AreEqual("AKKUYU_VALVE", model);
    }

    [TestMethod]
    public void GetObjectModel_ReturnsFanForAnSac()
    {
        var model = MechanismModelResolver.GetObjectModel("10SAC33AN007");
        Assert.AreEqual("AKKUYU_FAN", model);
    }
}

[TestClass]
public sealed class CsvExporterTests
{
    [TestMethod]
    public void CleanString_RemovesParentheses()
    {
        var cleaned = CsvExporter.CleanString("ABC(123)");
        Assert.AreEqual("ABC123", cleaned);
    }

    [TestMethod]
    public void ParseMechTriplet_ParsesThreeParts()
    {
        var (kks, num, bld) = CsvExporter.ParseMechTriplet("10SGK33AA007 18 10UBA");
        Assert.AreEqual("10SGK33AA007", kks);
        Assert.AreEqual("18", num);
        Assert.AreEqual("10UBA", bld);
    }

    [TestMethod]
    public void GetBlockFromBuildingOrKks_UsesBuildingFirst()
    {
        var block = CsvExporter.GetBlockFromBuildingOrKks("11UBA", "00CMM11");
        Assert.AreEqual("10", block);
    }

    [TestMethod]
    public void GetBlockFromBuildingOrKks_UsesKksFallback()
    {
        var block = CsvExporter.GetBlockFromBuildingOrKks(null, "00CMM11");
        Assert.AreEqual("00", block);
    }

    [TestMethod]
    public void BuildDeviceDescription_JoinsPartsWithSpaces()
    {
        var desc = CsvExporter.BuildDeviceDescription("(10UBB)", "10CMM11", "Тип");
        Assert.AreEqual("(10UBB) 10CMM11 Тип", desc);
    }

    [TestMethod]
    public void GetDistinctBuildingsWithExclusion_ReturnsOtherBuildings()
    {
        var kkss = new List<string>
        {
            "10SGK33AA007 18 10UBA",
            "10SGK33AA008 19 10UBB"
        };

        var result = CsvExporter.GetDistinctBuildingsWithExclusion(kkss, "10UBA");
        Assert.AreEqual("(10UBB)", result);
    }

    [TestMethod]
    public void GetDistinctBuildingsForBlockExport_ReturnsOtherBuildings()
    {
        var kkss = new List<string>
        {
            "10SGK33AA007 18 10UBA",
            "10SGK33AA008 19 10UBB"
        };

        var result = CsvExporter.GetDistinctBuildingsForBlockExport(kkss, "10UBA");
        Assert.AreEqual("(10UBB)", result);
    }
}

[TestClass]
public sealed class DplParserTests
{
    [TestMethod]
    public void SplitLine_RespectsQuotes()
    {
        var tokens = DplParser.SplitLine("A \"B C\" D");
        Assert.HasCount(3, tokens);
        Assert.AreEqual("A", tokens[0]);
        Assert.AreEqual("\"B C\"", tokens[1]);
        Assert.AreEqual("D", tokens[2]);
    }

    [TestMethod]
    public void ExtractData_ParsesCabinetCodeAndComment()
    {
        string line = "GmsDevice001 LANG:10027 \"Шлюз_ Пожар ППКиУП 10CMM11\"";
        var result = DplParser.ExtractData(line);

        Assert.AreEqual("GmsDevice001", result.dataPoint);
        Assert.AreEqual("10CMM11", result.cabinetCode);
        Assert.AreEqual("Шлюз_ Пожар ППКиУП", result.comment);
    }

    [TestMethod]
    public void GetMatchingLinesFromDplFile_FindsOnlyMatchingPatterns()
    {
        string tempFile = Path.GetTempFileName();
        try
        {
            var lines = new[]
            {
                "GmsDevice001 LANG:10027 \"Шлюз_ Пред_ тревога ППКиУП 10CMM11\"",
                "GmsDevice002 LANG:10027 \"Не тот текст\""
            };
            File.WriteAllLines(tempFile, lines);

            var matched = DplParser.GetMatchingLinesFromDplFile(tempFile);
            Assert.HasCount(1, matched);
        }
        finally
        {
            if (File.Exists(tempFile))
                File.Delete(tempFile);
        }
    }
}

[TestClass]
public sealed class ConversionParityTests
{
    [TestMethod]
    public void BlockExport_LinesMatchNormalConversion_ForExcelFile()
    {
        string excelPath = GetExcelPath();
        if (!File.Exists(excelPath))
        {
            Assert.Inconclusive($"Excel файл не найден: {excelPath}");
        }

        var cabinetData = ExcelParser.ReadCabinetData(excelPath, "Шкафы");
        var mechData = ExcelParser.ReadMechData(excelPath, "Мех-мы");

        var blockOutputs = BuildBlockOutputs(cabinetData, mechData);

        foreach (var cabinetEntry in cabinetData)
        {
            string cabinetKks = cabinetEntry.Key;
            if (!mechData.TryGetValue(cabinetKks, out var kkss) || kkss == null || kkss.Count == 0)
                continue;

            string building = cabinetEntry.Value.Building;
            string cabinetNm = CsvExporter.CleanString(cabinetKks);
            string ip = CsvExporter.CleanString(cabinetEntry.Value.IP);
            string type = CsvExporter.CleanString(cabinetEntry.Value.Type);

            string normal = CsvExporter.GenerateCsvContent(kkss, cabinetNm, ip, building, type).ToString();
            string prefix = $"{building}_{cabinetNm}";

            string block = blockOutputs[CsvExporter.GetBlockFromBuildingOrKks(building, cabinetKks)];

            var normalLines = ExtractCabinetLines(normal, prefix);
            var blockLines = ExtractCabinetLines(block, prefix);

            CollectionAssert.AreEqual(
                normalLines,
                blockLines,
                $"Несовпадение формата для шкафа {cabinetKks}"
            );
        }
    }

    private static string GetExcelPath()
    {
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        return Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "..", "СКУПЗ_АККУЮ_V12.17.xlsx"));
    }

    private static Dictionary<string, string> BuildBlockOutputs(
        Dictionary<string, (string IP, string Building, string Type, string VEKSH)> cabinetData,
        Dictionary<string, List<string>> mechData)
    {
        var byBlock = new Dictionary<string, List<string>>();
        foreach (var kv in cabinetData)
        {
            string cabKks = kv.Key;
            string building = kv.Value.Building;
            string block = CsvExporter.GetBlockFromBuildingOrKks(building, cabKks);
            if (block == null)
                continue;
            if (!byBlock.ContainsKey(block))
                byBlock[block] = new List<string>();
            byBlock[block].Add(cabKks);
        }

        string[] order = { "00", "10", "20", "30", "40" };
        var blocksOrdered = order.Where(b => byBlock.ContainsKey(b)).ToList();
        var result = new Dictionary<string, string>();

        foreach (var block in blocksOrdered)
        {
            var sb = new StringBuilder();
            CsvExporter.AppendStandardHeader(sb);
            CsvExporter.AppendDevicesHeader(sb);

            foreach (var cabKks in byBlock[block])
            {
                if (!mechData.TryGetValue(cabKks, out var kkss) || kkss == null || kkss.Count == 0)
                    continue;

                var cab = cabinetData[cabKks];
                string building = cab.Building;
                string cabinetNm = CsvExporter.CleanString(cabKks);
                string ip = CsvExporter.CleanString(cab.IP);
                string type = CsvExporter.CleanString(cab.Type);

                CsvExporter.AppendCabinetDeviceLine(sb, cabinetNm, ip, building, type, kkss);
            }

            sb.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            CsvExporter.AppendPointsHeader(sb);

            foreach (var cabKks in byBlock[block])
            {
                if (!mechData.TryGetValue(cabKks, out var kkss) || kkss == null || kkss.Count == 0)
                    continue;

                var cab = cabinetData[cabKks];
                string building = cab.Building;
                string cabinetNm = CsvExporter.CleanString(cabKks);
                string type = CsvExporter.CleanString(cab.Type);

                CsvExporter.AppendCabinetPoints(sb, cabinetNm, building, type, kkss);
            }

            sb.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            sb.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            sb.AppendLine(@"[VALUE_ASSIGNMENT],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            sb.AppendLine(@"#[ObjectPath],[Property Name],[Property Value],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");

            result[block] = sb.ToString();
        }

        return result;
    }

    private static List<string> ExtractCabinetLines(string content, string prefix)
    {
        return content
            .Split(new[] { "\r\n", "\n" }, StringSplitOptions.None)
            .Select(line => line?.TrimEnd() ?? string.Empty)
            .Where(line => line.StartsWith(prefix + ",", StringComparison.Ordinal))
            .ToList();
    }
}
