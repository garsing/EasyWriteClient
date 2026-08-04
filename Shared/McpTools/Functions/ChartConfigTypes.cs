using System.Collections.Generic;

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
    public int GapWidth { get; set; }
    public bool DataMarkers { get; set; }
    public int MarkerSize { get; set; }
    public double LineWeight { get; set; }
    public double YAxisMin { get; set; }
    public double YAxisMax { get; set; }
    public double YAxisMajorUnit { get; set; }
    public string DataLabelType { get; set; }
    public int Explosion { get; set; }
    public bool FillMissingValues { get; set; }
  }

  public class SeriesInfo
  {
    public string Name { get; set; }
    public string AxisY { get; set; }
    public string Color { get; set; }
    public string ChartType { get; set; }
    public bool? ShowDataLabels { get; set; }
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
    public int? TickLabelCount { get; set; }
    public int? TickLabelSpacing { get; set; }
    /// <summary>null 时按 Chart.Type 默认：line=false，column_clustered=true</summary>
    public bool? BetweenCategories { get; set; }
  }

  public class AxisYInfo
  {
    public string PrimaryTitle { get; set; }
    public string SecondaryTitle { get; set; }
  }

  public class ChartOperationResult
  {
    public bool Success { get; set; }
    public string Error { get; set; }
  }

  public sealed class ChartCreateResult : ChartOperationResult
  {
    public int ChartStart { get; set; }
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
