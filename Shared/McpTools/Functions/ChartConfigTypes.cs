using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
  public class ChartConfig
  {
    public ChartInfo Chart { get; set; }
  }

  public class ChartInfo
  {
    public string Type { get; set; }
    public string Title { get; set; }
    public List<SeriesInfo> SeriesCollection { get; set; }
    public string Theme { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
    public bool ShowDataLabels { get; set; }
    public string LegendPosition { get; set; }
    public string PlotColor { get; set; }
    public int TitleFontSize { get; set; }
    public bool TitleFontBold { get; set; }
    public string TitleFontColor { get; set; }
    public string Gridlines { get; set; }
    public AxisXInfo AxisX { get; set; }
    public AxisYInfo AxisY { get; set; }
    public AxisYInfo AxisY2 { get; set; }
    public int Overlap { get; set; }
    public string PlotBox { get; set; }
    public string PlotInside { get; set; }
    public string ChartAreaColor { get; set; }
    public int ChartStyle { get; set; }
    public string LegendFontColor { get; set; }
    public int GapWidth { get; set; }
    public bool DataMarkers { get; set; }
    public int MarkerSize { get; set; }
    public double LineWeight { get; set; }
    public double YAxisMin { get; set; }
    public double YAxisMax { get; set; }
    public double YAxisMajorUnit { get; set; }
    public bool? ShowValue { get; set; }
    public bool? ShowPercentage { get; set; }
    public int Explosion { get; set; }
    /// <summary>稿里写了 Explosion 才为 true。没写时套用保留图上已有的分离值。</summary>
    public bool HasExplosion { get; set; }
    public bool FillMissingValues { get; set; }
  }

  public class SeriesInfo
  {
    public string Name { get; set; }
    public string AxisY { get; set; }
    public string Color { get; set; }
    public string ChartType { get; set; }
    public bool? ShowDataLabels { get; set; }
    public bool? ShowValue { get; set; }
    public bool? ShowPercentage { get; set; }
    public string Line { get; set; }
    public string LineWeight { get; set; }
    public string Marker { get; set; }
    public string MarkerSize { get; set; }
    public string MarkerColor { get; set; }
    public string MarkerFill { get; set; }
    public string LabelPosition { get; set; }
    public string LabelFont { get; set; }
    public string LabelSize { get; set; }
    public string LabelColor { get; set; }
    public string FillGradient { get; set; }
    public string FillAngle { get; set; }
    public List<PointInfo> Points { get; set; }
  }

  public class PointInfo
  {
    public string X { get; set; }
    public double Y { get; set; }
  }

  public class AxisXInfo
  {
    public string Title { get; set; }
    public string Type { get; set; }
    public string Format { get; set; }
    public string Visible { get; set; }
    public string TickFont { get; set; }
    public string TickColor { get; set; }
    public string TickSize { get; set; }
    public string TickPosition { get; set; }
    public string MajorTick { get; set; }
    public string MinorTick { get; set; }
    public string Grid { get; set; }
    public int? TickLabelCount { get; set; }
    public int? TickLabelSpacing { get; set; }
    /// <summary>null 时按 Chart.Type 默认：line=false，column=true</summary>
    public bool? BetweenCategories { get; set; }
  }

  public class AxisYInfo
  {
    public string Title { get; set; }
    public string Visible { get; set; }
    public string TickFont { get; set; }
    public string TickColor { get; set; }
    public string TickSize { get; set; }
    public string TickPosition { get; set; }
    public string MajorTick { get; set; }
    public string MinorTick { get; set; }
    public string Grid { get; set; }
    public string Format { get; set; }
    public string Line { get; set; }
    public string LineWeight { get; set; }
    public string GridColor { get; set; }
  }

  public class ChartOperationResult
  {
    public bool Success { get; set; }
    public string Error { get; set; }
  }

  public sealed class ChartCreateResult : ChartOperationResult
  {
    public int ChartStart { get; set; }
    public Word.InlineShape InlineShape { get; set; }
  }

  public sealed class ChartValidationResult
  {
    public bool IsValid { get; set; }
    public string Error { get; set; }
  }

  public sealed class ChartXmlParseResult
  {
    public ChartConfig Config { get; set; }
    public string Error { get; set; }
    public string DataSource { get; set; }
    public string CsvFileName { get; set; }
  }
}
