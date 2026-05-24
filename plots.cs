using ScottPlot;
using ScottPlot.WinForms;
using ExcelDna.Integration;
using ScottPlot.Testing;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;

public class PlotWindow : Form
{
    public ScottPlot.WinForms.FormsPlot PlotControl { get; private set; }

    public bool lcked=false;
    public string wb;
    public string ws;

    public ScottPlot.Plottables.Scatter nodescatter;
    public ScottPlot.Plottables.Scatter elementscatter;

    public ScottPlot.Plottables.Scatter effectscatter;

    public string rng;
    public PlotWindow()
    {
        Text = "Interactive Plot";
        Width = 900;
        Height = 700;

        PlotControl = new ScottPlot.WinForms.FormsPlot
        {
            Dock = DockStyle.Fill
        };
        Controls.Add(PlotControl);
        
        PlotControl.Menu.Add("Toggle Auto rescale on refresh", (plt) =>
        {
            lcked = !lcked;
        });

        
    }

    public void UpdatePlot(string rng)
    {
        var plt = PlotControl.Plot;
        plt.Clear();
        string frm=((string)ExcelDnaUtil.Application.Workbooks[wb].Worksheets[ws].Range[rng].Formula2).TrimEnd(")".ToCharArray());
        string[] splithold=frm.Split("graphvoid(");
        splithold=splithold[splithold.Length-1].Split(",");
        List<object[,]> cellvalues=new List<object[,]>();
        object[,] holderarray1;
        object[,] holderarray2;
        object valholder;
        Array arr;
        for (int i=0; i<splithold.Length; i++)
        {
            valholder=ExcelDnaUtil.Application.Workbooks[wb].Worksheets[ws].Range[splithold[i]].Value;
            if (valholder is not object[,])
            {
                arr=Array.CreateInstance(typeof(string), new int[] {1,1}, new int[]{1,1});
                arr.SetValue(valholder.ToString(),1,1);
                valholder=arr;
            }
            holderarray1=(object[,])valholder;
            holderarray2=new object[holderarray1.GetLength(0),holderarray1.GetLength(1)];
            for (int j=0; j<holderarray1.GetLength(0); j++)
            {
                for (int k=0; k<holderarray1.GetLength(1); k++)
                {
                    holderarray2[j,k]=holderarray1[j+1,k+1]; 
                }
            }
            cellvalues.Add(holderarray2);
        }
        try
        {
            object[,] effectcoords=TESTEXDNA.graphingtablefunctions.grapheffects(TESTEXDNA.interfacefunctions.filterarrayempties(cellvalues[2]),TESTEXDNA.interfacefunctions.filterarrayempties(cellvalues[3]),cellvalues[0],cellvalues[1],cellvalues[4],cellvalues[5]);
            drawlines(effectcoords, "effects");
        }
        catch
        {
            
        }
        
        object[,] elementcoords=TESTEXDNA.graphingtablefunctions.elementcoords(TESTEXDNA.interfacefunctions.filterarrayempties(cellvalues[2]),TESTEXDNA.interfacefunctions.filterarrayempties(cellvalues[3]));
        drawlines(elementcoords, "elements");
        object[,] nodecoords=TESTEXDNA.graphingtablefunctions.nodecoords(TESTEXDNA.interfacefunctions.filterarrayempties(cellvalues[2]));
        drawlines(nodecoords, "nodes");
        if (!lcked)
        {
            PlotControl.Plot.Axes.AutoScale();
        }
        PlotControl.Refresh();
    }
    public void drawlines(object[,] pts, object kwd)
    {
        (double[] x, double[] y)= extractxy(pts);
        var plt = PlotControl.Plot;
        switch (kwd)
        {
            case "nodes":
                nodescatter = plt.Add.Scatter(x, y);
                nodescatter.MarkerShape = MarkerShape.FilledCircle;
                nodescatter.MarkerSize = 10;
                nodescatter.LineWidth = 0;
                nodescatter.Color = Colors.Red;
                nodescatter.LegendText = "Nodes";
                break;
            case "elements":
                elementscatter = plt.Add.Scatter(x, y);
                elementscatter.MarkerSize = 0;
                elementscatter.LineWidth = 5;
                elementscatter.Color = Colors.Black;
                elementscatter.LegendText = "Elements";
                break;
            case "effects":
                effectscatter = plt.Add.Scatter(x, y);
                effectscatter.MarkerSize = 5;
                effectscatter.LineWidth = 2;
                effectscatter.Color = Colors.Blue;
                effectscatter.LegendText = "Effects";
                break;
            case null:
                break;
            
            default:
                throw new ArgumentException("Invalid keyword argument for drawlines method");
        }
    }
    public (double[] x, double[] y) extractxy(object[,] pts)
    {
        double[] x = new double[pts.GetLength(0)];
        double[] y = new double[pts.GetLength(0)];
        for (int i=0; i<pts.GetLength(0); i++)
        {
            if (pts[i,0] is ExcelError.ExcelErrorNA)
             {
                x[i]=double.NaN;
                y[i]=double.NaN;
             }
             else
             {
                x[i]=(double)pts[i,0];
                y[i]=(double)pts[i,1];
             }
            
        }
        return (x,y);
    }
}

public static class PlotManager
{
    private static PlotWindow _window;
    
    public static void ShowPlot(string wb, string ws, string rng)
    {
        if (_window == null || _window.IsDisposed)
        {   
            _window = new PlotWindow();
            _window.wb=wb;
            _window.ws=ws;
            _window.rng=rng;
            _window.Show();
            _window.BringToFront();
        }
        else
        {
            return;
        }
        if (rng != null)
        {
            ScottPlot.AxisRules.SquareZoomOut rule = new(
            xAxis: _window.PlotControl.Plot.Axes.Bottom,
            yAxis: _window.PlotControl.Plot.Axes.Left);

            _window.PlotControl.Plot.Axes.Rules.Clear();
            _window.PlotControl.Plot.Axes.Rules.Add(rule);
            _window.PlotControl.Refresh();
            UpdatePlotcontroller(rng);
        }
    }

    public static void UpdatePlotcontroller(string rng)
    {
        if (!(_window == null || _window.IsDisposed) )
        {
            _window?.UpdatePlot(rng);
        }
    }
}