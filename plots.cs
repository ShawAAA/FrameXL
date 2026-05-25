using ScottPlot;
using ScottPlot.WinForms;
using ExcelDna.Integration;
using ScottPlot.Testing;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using ScottPlot.Plottables;
using System.Security.Cryptography.X509Certificates;
using System.Drawing.Drawing2D;

public class PlotWindow : Form
{
    public ScottPlot.WinForms.FormsPlot PlotControl { get; private set; }

    public bool lcked=false;
    public string wb;
    public string ws;
    public ScottPlot.Plottables.Scatter nodescatter;
    public ScottPlot.Plottables.Scatter elementscatter;
    public List<object[,]> cellvalues;
    public ScottPlot.Plottables.Scatter effectscatter;
    public object[,] effectcoords;
    public object[,] effectcoordszeroed;
    public object tab2;
    public Tooltip ntooltip;
    public Tooltip etooltip;
    public Tooltip gtooltip;
    public string rng;
    public bool nodeorelselected=false;
    public bool effectselected=false;

    public bool drageffct=false;
    public PlotWindow()
    {
        Text = "FrameXL Interactive Plot";
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
        effectscatter=null;
        nodescatter=null;
        elementscatter=null;
        string frm=((string)ExcelDnaUtil.Application.Workbooks[wb].Worksheets[ws].Range[rng].Formula2).TrimEnd(")".ToCharArray());
        string[] splithold=frm.Split("graphvoid(");
        splithold=splithold[splithold.Length-1].Split(",");
        cellvalues=new List<object[,]>();
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
            (effectcoords,tab2) =TESTEXDNA.graphingtablefunctions.grapheffects(TESTEXDNA.interfacefunctions.filterarrayempties(cellvalues[2]),TESTEXDNA.interfacefunctions.filterarrayempties(cellvalues[3]),cellvalues[0],cellvalues[1],cellvalues[4],cellvalues[5]);
            string actionsdisps=cellvalues[1][0,1].ToString()??"";
            string lgnd;
            if (actionsdisps.ToLower() != "displacements")
            {
                lgnd="LC"+(cellvalues[1][0,0].ToString()??"")+" "+(cellvalues[1][0,2].ToString()??"")+" "+actionsdisps+" Local Direction "+(cellvalues[1][0,3].ToString()??"");
            }
            else
            {
                lgnd="LC"+(cellvalues[1][0,0].ToString()??"")+" "+(cellvalues[1][0,2].ToString()??"")+" "+actionsdisps+" Max/Min "+(cellvalues[1][0,3].ToString()??"")+" Dirn. Permutations";
            }
            drawlines(effectcoords, "effects");
            effectscatter.LegendText=lgnd;
            PlotControl.Plot.Title(lgnd);
        }
        catch
        {
            if (effectscatter != null)
            {
                plt.Remove(effectscatter);
                effectselected=false;
                effectscatter=null;
            }
        }
        try
        {
            object[,] elementcoords=TESTEXDNA.graphingtablefunctions.elementcoords(TESTEXDNA.interfacefunctions.filterarrayempties(cellvalues[2]),TESTEXDNA.interfacefunctions.filterarrayempties(cellvalues[3]));
            drawlines(elementcoords, "elements");
        }
        catch
        {
            if (elementscatter != null)
            {
                plt.Remove(elementscatter);
                nodeorelselected=false;
                elementscatter=null;
            }
        }
        try
        {
            object[,] nodecoords=TESTEXDNA.graphingtablefunctions.nodecoords(TESTEXDNA.interfacefunctions.filterarrayempties(cellvalues[2]));
            drawlines(nodecoords, "nodes");
        }
        catch
        {
            if (nodescatter != null)
            {
                plt.Remove(nodescatter);
                nodeorelselected=false;
                nodescatter=null;
            }
        }
        
        if (!lcked)
        {
            PlotControl.Plot.Axes.AutoScale();
        }
        ntooltip=PlotControl.Plot.Add.Tooltip(new Coordinates(0, 0), "Hover over a point", new Coordinates(0, 0));
        etooltip=PlotControl.Plot.Add.Tooltip(new Coordinates(0, 0), "Hover over a point", new Coordinates(0, 0));
        gtooltip=PlotControl.Plot.Add.Tooltip(new Coordinates(0, 0), "Hover over a point", new Coordinates(0, 0));
        ntooltip.IsVisible=false;
        etooltip.IsVisible=false;
        gtooltip.IsVisible=false;
        PlotControl.Refresh();
    }
    public void nttips(Pixel mousePixel)
    {
        Coordinates mouseLocation = PlotControl.Plot.GetCoordinates(mousePixel);
        DataPoint nearest = nodescatter.Data.GetNearest(mouseLocation, PlotControl.Plot.LastRender);
        if (nearest.IsReal)
        {
            ntooltip.LabelText = $"Node: {nearest.Index+1:F0}\nx: {nearest.Coordinates.X:F2}\ny: {nearest.Coordinates.Y:F2}";
            ntooltip.TipLocation = nearest.Coordinates;
            ntooltip.LabelLocation= nearest.Coordinates;
            ntooltip.LabelBackgroundColor = Colors.White.WithAlpha(1); 
            ntooltip.FillColor = Colors.White.WithAlpha(1); 
            ntooltip.LineColor = Colors.Transparent;
            ntooltip.LabelOffsetX=40;
            ntooltip.LabelOffsetY=-40;
            nodeorelselected=true;
            ntooltip.IsVisible = true;
        }
        else
        {
            nodeorelselected=false;
            ntooltip.IsVisible = false;
        }
        PlotControl.Refresh();
    }
    public void ettips(Pixel mousePixel)
    {
        Coordinates mouseLocation = PlotControl.Plot.GetCoordinates(mousePixel);
        DataPoint nearest = elementscatter.Data.GetNearest(mouseLocation, PlotControl.Plot.LastRender);
        double close;
        List<(int,Coordinates)> closemax=new List<(int,Coordinates)>();
        if (nearest.IsReal)
        {
            close=nearest.Coordinates.Distance(PlotControl.Plot.GetCoordinates(mousePixel));
            int indx=1;
            foreach (Coordinates coord in elementscatter.Data.GetScatterPoints())
            {
                if (close == coord.Distance(PlotControl.Plot.GetCoordinates(mousePixel)))
                {
                    closemax.Add((indx,coord));
                }
                indx=indx+1;
            }
            string elab="";
            for (int i = 0; i < closemax.Count(); i++)
            {
                elab=elab+$"Element: {Math.Floor((double)closemax[i].Item1/3)+1:F0}\nEnd: {closemax[i].Item1%3:F0}\nx: {closemax[i].Item2.X:F2}\ny: {closemax[i].Item2.Y:F2}\n";
            }
            elab.TrimEnd();
            etooltip.LabelText = elab;
            etooltip.TipLocation = nearest.Coordinates;
            etooltip.LabelLocation= nearest.Coordinates;
            etooltip.LabelBackgroundColor = Colors.White.WithAlpha(1); 
            etooltip.FillColor = Colors.White.WithAlpha(1); 
            etooltip.LineColor = Colors.Transparent;
            etooltip.LabelOffsetX=40;
            etooltip.LabelOffsetY=25+35*closemax.Count;
            nodeorelselected=true;
            etooltip.IsVisible = true;
        }
        else
        {
            nodeorelselected=false;
            etooltip.IsVisible = false;
        }
        PlotControl.Refresh();
    }
    public void gttips(Pixel mousePixel)
    {
        Coordinates mouseLocation = PlotControl.Plot.GetCoordinates(mousePixel);
        DataPoint nearest = effectscatter.Data.GetNearest(mouseLocation, PlotControl.Plot.LastRender);
        if (nearest.IsReal)
        {
            gtooltip.LabelText = $"Effect: {nearest.Index+1:F0}\nx: {nearest.Coordinates.X:F2}\ny: {nearest.Coordinates.Y:F2}";
            gtooltip.TipLocation = nearest.Coordinates;
            gtooltip.LabelLocation= nearest.Coordinates;
            gtooltip.LabelBackgroundColor = Colors.White.WithAlpha(1); 
            gtooltip.FillColor = Colors.White.WithAlpha(1); 
            gtooltip.LineColor = Colors.Transparent;
            gtooltip.LabelOffsetX=40;
            gtooltip.LabelOffsetY=-40;
            effectselected=true;
            gtooltip.IsVisible = true;
        }
        else
        {
            effectselected=false;
            gtooltip.IsVisible = false;
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
            _window.PlotControl.Plot.XLabel("x (m)");
             _window.PlotControl.Plot.YLabel("y (m)");
            
            _window.PlotControl.Refresh();
            UpdatePlotcontroller(rng);
            
            _window.PlotControl.MouseMove += (s, e) =>
            {
                Pixel mousePixel = new(e.Location.X, e.Location.Y);
                if (!_window.drageffct)
                {
                    if (_window.nodescatter!=null)
                    {
                        _window.nttips(mousePixel);
                    }
                    if (_window.elementscatter!=null)
                    {
                        _window.ettips(mousePixel);
                    }
                    if (!_window.nodeorelselected && _window.effectscatter!=null)
                    {
                        _window.gttips(mousePixel);
                    }
                }
                else
                {
                    
                }
                
            };
            _window.PlotControl.MouseDown += (object? sender, MouseEventArgs e) =>
            {
                if (_window.effectselected)
                {
                    _window.drageffct=true;
                    object[,] zrorequest=(object[,])_window.cellvalues[1].Clone();
                    zrorequest[0,4]=0;
                    _window.effectcoordszeroed=TESTEXDNA.graphingtablefunctions.grapheffects(TESTEXDNA.interfacefunctions.filterarrayempties(_window.cellvalues[2]),TESTEXDNA.interfacefunctions.filterarrayempties(_window.cellvalues[3]),_window.cellvalues[0],zrorequest,_window.cellvalues[4],_window.cellvalues[5]).Item1;
                    _window.PlotControl.UserInputProcessor.Disable();
                    _window.gtooltip.IsVisible = false;
                    _window.PlotControl.Refresh();
                }
                
            };
            _window.PlotControl.MouseUp += (object? sender, MouseEventArgs e) =>
            {
                if (_window.effectselected)
                {
                    _window.drageffct=false;
                    _window.PlotControl.UserInputProcessor.Enable();
                    _window.gtooltip.IsVisible = true;
                }
            };

        }
    }

    public static void UpdatePlotcontroller(string rng)
    {
        if (!(_window == null || _window.IsDisposed) )
        {
            _window?.UpdatePlot(rng);
        }
    }
    public static void kill()
    {
        if (!(_window == null || _window.IsDisposed) )
        {
            _window.Dispose();
        }
    }
}