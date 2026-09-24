"""Audit theme coverage; optionally compile production XAML and theme code.

The compilation harness stubs unrelated SSMS/database event handlers. It checks
WPF markup and SDK/library API compatibility, not SSMS hosting or rendering.
"""
import argparse
from collections import Counter
from pathlib import Path
import re
import subprocess
import tempfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

ROOT = Path(__file__).resolve().parents[2] / "AxialSqlTools"
X = "{http://schemas.microsoft.com/winfx/2006/xaml}"
WPF = "{http://schemas.microsoft.com/winfx/2006/xaml/presentation}"


def audit():
    pages = sorted(ROOT.rglob("*.xaml"))
    shared = ET.parse(ROOT / "Themes/SharedToolWindowTheme.xaml").getroot()
    keys = [element.get(X + "Key") or element.get("TargetType") for element in shared]
    duplicates = [key for key, count in Counter(keys).items() if key and count > 1]
    assert not duplicates, f"Duplicate shared resource keys: {duplicates}"
    resources = {element.get(X + "Key") for element in shared}
    views = 0
    for page in pages:
        root = ET.parse(page).getroot()
        if root.get(X + "Class"):
            views += 1
            assert "SharedToolWindowTheme.xaml" in page.read_text(encoding="utf-8-sig"), page
            for prop in ("Background", "Foreground"):
                assert root.get(prop, "").startswith("{DynamicResource AxialTheme"), (page, prop)
            code = page.with_suffix(".xaml.cs").read_text(encoding="utf-8-sig")
            assert "new ToolWindowThemeController(" in code, f"No live theme subscription: {page}"
        for element in root.iter():
            for attr, value in element.attrib.items():
                for key in re.findall(r"\{DynamicResource (AxialTheme\w+)\}", value):
                    assert key in resources, f"Missing theme brush {key}: {page}"
                prop = element.get("Property") if attr == "Value" else attr
                if page.parent.name == "Themes":
                    continue  # The shared dictionary contains design-time palette defaults.
                if prop not in {"Background", "Foreground", "BorderBrush", "Fill", "Stroke", "RowBackground",
                                "AlternatingRowBackground", "HorizontalGridLinesBrush", "VerticalGridLinesBrush"}:
                    continue
                if value.startswith("{") or value == "Transparent":
                    continue
                # This swatch represents the user's connection color, not UI chrome.
                assert element.get(X + "Name") == "NewRuleColorPreview" and prop == "Background", (page, prop, value)
    for source in ROOT.rglob("*.cs"):
        assert not re.search(r"new Setter\([^;\n]+(?:Brushes\.|Colors\.|SystemColors\.)", source.read_text(encoding="utf-8-sig")), f"A generated style captures a fixed color: {source}"
    print(f"PASS: {views} views have shared resources, surface colors and live theme subscriptions; {len(pages)} XAML files parse.")
    return pages


def extract_method(source, name):
    """Copy the production method body verbatim into an API-compatibility harness."""
    match = re.search(r"^\s*(?:private|internal) (?:static )?[^\n]+\b" + name + r"\(", source, re.M)
    assert match, name
    opening = source.index("{", match.start())
    depth = 1
    end = opening + 1
    while depth:
        if source[end] == "{": depth += 1
        if source[end] == "}": depth -= 1
        end += 1
    return source[match.start():end]


def build(pages, dotnet, framework):
    with tempfile.TemporaryDirectory(prefix="axial-theme-validation-") as temporary:
        destination = Path(temporary)
        classes = []
        for page in pages:
            root = ET.parse(page).getroot()
            qualified_name = root.get(X + "Class")
            if not qualified_name:
                continue
            namespace, name = qualified_name.rsplit(".", 1)
            source = page.with_suffix(".xaml.cs").read_text(encoding="utf-8-sig")
            methods = set(re.findall(r"(?:private|internal|public|protected)\s+(?:async\s+|static\s+|override\s+)*[\w.<>,?]+\s+(\w+)\s*\(", source))
            handlers = {value for element in root.iter() for value in element.attrib.values() if value in methods}
            base = {"Window": "System.Windows.Window", "UserControl": "System.Windows.Controls.UserControl",
                    "DialogWindow": "Microsoft.VisualStudio.PlatformUI.DialogWindow"}[root.tag.split("}")[-1]]
            code = f"namespace {namespace} {{ public partial class {name} : {base} {{\n"
            code += "\n".join(f"private void {handler}(object sender, System.EventArgs e) {{ }}" for handler in sorted(handlers))
            if name == "ComparisonEndpointControl":
                code += "\npublic event System.EventHandler EndpointChanged;"
            if name == "QuickSearchWindowControl":
                code += "\nprivate TextMarkerService textMarkerService;"
                code += extract_method(source, "ApplyThemeBrushResources")
            if name == "HealthDashboard_ServerControl":
                code += "\n" + "\n".join(re.findall(r"private (?:OxyColor|Brush) _\w+[^;]*;", source))
                code += "\n" + re.search(r"private readonly ConditionalWeakTable<[^;]+;", source).group(0)
                for method in ("ApplyThemeBrushResources", "ApplyPlotTheme", "CapturePlotColors", "ReadablePlotColor", "ToOxyColor", "GetThemeBrush"):
                    code += extract_method(source, method)
            if name == "SqlServerBuildsWindowControl":
                code += "\nprivate void CopyUrlToClipboard(string url) { } private void OpenUrl(string url) { }"
                for method in ("CreateColumnText", "CreateHyperlink"):
                    code += extract_method(source, method)
            if name == "PivotGridWindowControl":
                for method in ("CreateCellStyle", "CreateTextStyle"):
                    code += extract_method(source, method)
            classes.append(code + "\n} }")
        classes.append("namespace AxialSqlTools { static class SQLBuilds { public static bool TryHttpUrl(string url, out string normalized) { normalized = url; return true; } } }")
        imports = """using System;
using System.Data;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using System.Xml;
using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Highlighting.Xshd;
using OxyColor = OxyPlot.OxyColor;
using PlotModel = OxyPlot.PlotModel;
using OxyPlot.Series;
"""
        (destination / "ViewStubs.cs").write_text(imports + "\n".join(classes), encoding="utf-8")
        project = f"""<Project Sdk="Microsoft.NET.Sdk">
<PropertyGroup><TargetFramework>{framework}</TargetFramework><AssemblyName>AxialSqlTools</AssemblyName>
<UseWPF>true</UseWPF><EnableWindowsTargeting>true</EnableWindowsTargeting>
<LangVersion>latest</LangVersion><EnableDefaultCompileItems>false</EnableDefaultCompileItems>
<EnableDefaultPageItems>false</EnableDefaultPageItems><RunAnalyzersDuringBuild>false</RunAnalyzersDuringBuild>
<NoWarn>CS0067;CS0649</NoWarn></PropertyGroup><ItemGroup>
"""
        project += "\n".join(f'<Page Include="{escape(str(page))}" Link="{escape(str(page.relative_to(ROOT)))}" />' for page in pages)
        project += '<Compile Include="ViewStubs.cs" />'
        for file in ("Modules/ToolWindowThemeSupport.cs", "QuickSearch/TextMarkerService.cs"):
            project += f'<Compile Include="{escape(str(ROOT / file))}" />'
        production = ET.parse(ROOT / "AxialSqlTools.csproj").getroot()
        ns = {"m": "http://schemas.microsoft.com/developer/msbuild/2003"}
        for package in ("AvalonEdit", "OxyPlot.Wpf"):
            reference = production.find(f'.//m:PackageReference[@Include="{package}"]/m:Version', ns)
            project += f'<PackageReference Include="{package}" Version="{reference.text}" />'
        project += '<PackageReference Include="Microsoft.VisualStudio.Shell.15.0" Version="17.14.40264" /></ItemGroup></Project>'
        (destination / "ThemeValidation.csproj").write_text(project, encoding="utf-8")
        subprocess.run([dotnet, "build", str(destination / "ThemeValidation.csproj"), "-c", "Release", "--verbosity", "quiet"], check=True)
        print(f"PASS: compiled all {len(pages)} production XAML files, shared theme code, pivot styles, chart theme methods and SQL preview theme methods ({framework}).")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--build", action="store_true", help="Also compile WPF markup and changed theme methods using the .NET 8 SDK")
    parser.add_argument("--dotnet", default="dotnet")
    parser.add_argument("--framework", default="net48")
    options = parser.parse_args()
    pages = audit()
    if options.build:
        build(pages, options.dotnet, options.framework)
