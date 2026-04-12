using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.VisualStyles;
using System.Drawing.Imaging;

namespace TESTEXDNA
{
  [ComVisible(true)]
  public class RibbonController : ExcelRibbon
  {
    public override string GetCustomUI(string RibbonID)
    {
      return @"
      <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
      <ribbon>
        <tabs>
          <tab id='tab1' label='FrameXL'>
            <group id='group1' label='Setup'>
              <button id='button1' label='New Structure - Pure Text Output' onAction='SetupNewStruct' tag='text'/> 
              <button id='button2' label='New Structure - Graphical/Table Output' onAction='SetupNewStruct' tag='simple'/>
            </group >
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
    }

public void SetupNewStruct(ExcelDna.Integration.CustomUI.IRibbonControl control)
    {
      //MessageBox.Show("Hello from control " + control.Id);
      object tg=control.Tag;
      dynamic xlApp = ExcelDnaUtil.Application;
      String wb = xlApp.ActiveWorkbook.name;
      String ws = xlApp.ActiveSheet.name;
      String cell = xlApp.ActiveCell.offset(1).address;
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset(-1).value="FrameXL V0.1 - 03/04/2-26. Created by A. Shaw. Refer notes for guidance.";
      int offs=0;

      string[] nodes1 = new string[] { "x Coord.(m)", "y Coord. (m)", "x fixity","y fixity","Rotation fixity"};
      object[] nodes2 = new object[] { 0, 0, "True","True","True"};
      object[] nodes3 = new object[] { 1, 0, "False","False","False"};
      createinputblock(xlApp, wb, ws, cell, offs, "Node Inputs", nodes1, 2, nodes2,nodes3);
      Comment cmt=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,2].addcomment("Fixities as True/False or with stiffness N/m or N/rad. Input spring curves with '-disp/rot,-force/moment;...,...;disp/rot,force/moment'. 'disp/rot' coordinates are in global reference system.");
      cmt.Shape.TextFrame.Characters().Font.Bold = false;
      cmt.Shape.TextFrame.AutoSize = true;
      offs=offs+6;
      
      string[] element1 = new string[] { "Node 1 Index", "Node 2 Index", "EA (N)","EI (Nm2)","End 1 Axial Release","End 1 Shear Release","End 1 Rotational Release","End 2 Axial Release","End 2 Shear Release","End 2 Rotational Release"};
      object[] element2 = new object[] { 1, 2, 1000,1000,false,false,false,false,false,false};
      createinputblock(xlApp, wb, ws, cell, offs, "Element Inputs", element1, 1, element2);
      offs=offs+5;

      string[] nodeloads1 = new string[] { "Load Case Number", "Node Applied", "Load Direction","Load Magnitude (N or Nm)"};
      object[] nodeloads2 = new object[] { 1, 2, "x",1};
      object[] nodeloads3 = new object[] { 1, 2, "y",1};
      createinputblock(xlApp, wb, ws, cell, offs, "Node Load Inputs", nodeloads1, 2, nodeloads2,nodeloads3);
      cmt=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs].addcomment("Load applied in global coordinates.");
      cmt.Shape.TextFrame.Characters().Font.Bold = false;
      cmt.Shape.TextFrame.AutoSize = true;
      cmt=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,1].addcomment("Node lists can be given as 'All', a list of indexes '1,2,3' or a range '1to3,5to-1'.");
      cmt.Shape.TextFrame.Characters().Font.Bold = false;
      cmt.Shape.TextFrame.AutoSize = true;
      cmt=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,3].addcomment("Magnitude or equation for magnitude. Equations are input as '#index+lc+2*x+y+2', or similar, with properties derived from the assigned node. ");
      cmt.Shape.TextFrame.Characters().Font.Bold = false;
      cmt.Shape.TextFrame.AutoSize = true;
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 2, 2].validation.add(3, 3, 3, "x,y,zz", 0);
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 3, 2].validation.add(3, 3, 3, "x,y,zz", 0);
      offs=offs+6;

      string[] elementloads1 = new string[] { "Load Case Number", "Element Applied", "Point/Patch","Local/Global/Global Projected","Direction","Point Load Position / LHS of Patch (%)","Point Load Magnitude (N), LHS Patch Magnitude (N/m)","RHS of Patch (%)","RHS Patch Magnitude (N/m)"};
      object[] elementloads2 = new object[] { 3, 1, "Point Load","Global Projected","y",0.5,-1};
      createinputblock(xlApp, wb, ws, cell, offs, "Beam Load Inputs", elementloads1, 1, elementloads2);
      cmt=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,3].addcomment("Local applied x/axial, y/vertical and zz/rotational. Global Projected Applied across the projected length of the element (N/A for rotation).");
      cmt.Shape.TextFrame.Characters().Font.Bold = false;
      cmt.Shape.TextFrame.AutoSize = true;
      cmt=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,1].addcomment("Element lists can be given as 'All', a list of indexes '1,2,3' or a range '1to3,5to-1'.");
      cmt.Shape.TextFrame.Characters().Font.Bold = false;
      cmt.Shape.TextFrame.AutoSize = true;
      cmt=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,6].addcomment("Magnitude or equation for magnitude. Equations are input as '#index+lc+2*x+y+2', or similar, with properties derived from the position at which the load 'turns-on' evaluated at either the point load or either side of the path.");
      cmt.Shape.TextFrame.Characters().Font.Bold = false;
      cmt.Shape.TextFrame.AutoSize = true;
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 2, 2].validation.add(3, 3, 3, "Point Load,Patch Load", 0);
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 3, 2].validation.add(3, 3, 3, "Point Load,Patch Load", 0);
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 2, 3].validation.add(3, 3, 3, "Local,Global,Global Projected", 0);
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 3, 3].validation.add(3, 3, 3, "Local,Global,Global Projected", 0);
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 2, 4].validation.add(3, 3, 3, "x,y,zz", 0);
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 3, 4].validation.add(3, 3, 3, "x,y,zz", 0);
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 2, 5].numberformat="0.00%";
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 3, 5].numberformat="0.00%";
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 2, 7].numberformat="0.00%";
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 3, 7].numberformat="0.00%";
      offs=offs+5;

      if ((string)tg == "text")
      {
        string[] extraction1 = new string[] { "Load Case List", "Actions/Displacements", "Nodes/Elements","Element/Node Index List"};
        object[] extraction2 = new object[] { "All", "Actions", "Elements","1"};
        createinputblock(xlApp, wb, ws, cell, offs, "Extracted Results", extraction1, 1, extraction2);
        cmt=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs].addcomment("Load case/node/element lists can be given as 'All', a list of indexes '1,2,3' or a range '1to3,5to6'. Node 'actions' are the reaction forces only, while element actions are the internal actions. Element actions/displacements are in local coords while node actions/displacements are in global coords.");
        cmt.Shape.TextFrame.Characters().Font.Bold = false;
        cmt.Shape.TextFrame.AutoSize = true;
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 2, 1].validation.add(3, 3, 3, "Actions,Displacements", 0);
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 3, 1].validation.add(3, 3, 3, "Actions,Displacements", 0);
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 2, 2].validation.add(3, 3, 3, "Nodes,Elements", 0);
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 3, 2].validation.add(3, 3, 3, "Nodes,Elements", 0);
        offs=offs+5;
        xlApp.workbooks[wb].worksheets[ws].range[cell].resize[20,10].Columns.Autofit();
      
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs].formula2 = "=structuralanalysis(" + xlApp.range(cell).offset(2, 0).resize(3,5).address +","+ xlApp.range(cell).offset(8, 0).resize(2,10).address+","+ xlApp.range(cell).offset(13, 0).resize(3,4).address+","+ xlApp.range(cell).offset(19, 0).resize(2,9).address +","+ xlApp.range(cell).offset(offs-3, 0).resize(2,4).address+ ")";
      }
      else if ((string)tg == "simple")
      {
        string[] graphresults1 = new string[] { "Graph LC", "Graph Actions/Displacements", "Graph Elements/Nodes","Graph Action Direction","Graph scale factor","Tabularise Actions/Displacements","Tabularise Elements/Nodes","Tabularise Load Case List","Tabularise Element/Node Index List"};
        object[] graphresults2 = new object[] { 1, "Actions", "Elements","x",100,"Actions","Elements","All","All"};
        createinputblock(xlApp, wb, ws, cell, offs, "Graph/Table Controls", graphresults1, 0, graphresults2);
        cmt=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs].addcomment("Graphing/Tabularisation formulas can be duplicated to extract multiple sets of results at the same time. Load case/node/element lists can be given as 'All', a list of indexes '1,2,3' or a range '1to3,5to6'. Node 'actions' are the reaction forces only, while element actions are the internal actions. Element actions/displacements are in local coords while node actions/displacements are in global coords. ");
        cmt.Shape.TextFrame.Characters().Font.Bold = false;
        cmt.Shape.TextFrame.AutoSize = true;
        cmt=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,3].addcomment("Graphs either local beam x, y or zz actions (axial, shear, bending respectively). Not used for displacements as displacements are directly rendered.");
        cmt.Shape.TextFrame.Characters().Font.Bold = false;
        cmt.Shape.TextFrame.AutoSize = true;
        cmt=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,4].addcomment("Scale factor for displacements, arbitrary graphing factor for actions");
        cmt.Shape.TextFrame.Characters().Font.Bold = false;
        cmt.Shape.TextFrame.AutoSize = true;
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 2, 1].validation.add(3, 3, 3, "Actions,Displacements,Loads", 0);
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 2, 2].validation.add(3, 3, 3, "Nodes,Elements", 0);
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 2, 3].validation.add(3, 3, 3, "x,y,zz", 0);
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 2, 5].validation.add(3, 3, 3, "Actions,Displacements,Loads", 0);
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 2, 6].validation.add(3, 3, 3, "Nodes,Elements", 0);
        string requestsaddress=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 2, 0].address;
        offs=offs+4;
        xlApp.workbooks[wb].worksheets[ws].range[cell].resize[20,10].Columns.Autofit();
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs].formula2 = "=structuralanalysissimple(" + xlApp.range(cell).offset(2, 0).resize(3,5).address +","+ xlApp.range(cell).offset(8, 0).resize(2,10).address+","+ xlApp.range(cell).offset(13, 0).resize(3,4).address+","+ xlApp.range(cell).offset(19, 0).resize(2,9).address + ")";
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs].resize(5,6).HorizontalAlignment=5;
        string resultsblockad=(string)xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs].address();
        offs=offs+5;
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs,6].value="Results Table";
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,6].formula2="=structuralanalysistabularise("+resultsblockad+"#,"+ xlApp.range(cell).offset(offs-7, 0).resize(1,9).address+","+xlApp.range(cell).offset(2, 0).resize(3,5).address+","+xlApp.range(cell).offset(8, 0).resize(2,10).address+","+ xlApp.range(cell).offset(13, 0).resize(3,4).address+","+ xlApp.range(cell).offset(19, 0).resize(2,9).address+")";
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,6].resize[1,7].Columns.Autofit();
        double lft = xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs-5].left;
        double top = xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs-5].top;
        double rght =xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs-5,6].left;
        ChartObject chartob = xlApp.workbooks[wb].worksheets[ws].ChartObjects.Add(lft, top, rght-lft, 0.9*(rght-lft));
        Chart chrt = chartob.Chart;
        chrt.PlotVisibleOnly = false;
        chrt.HasTitle = true;

        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs].formula2="=\"LC\" & "+requestsaddress + "& \", \" & " + xlApp.range(requestsaddress).offset(0,2).address + "& \", \" & "+ xlApp.range(requestsaddress).offset(0,1).address+ " & if("+xlApp.range(requestsaddress).offset(0,1).address +"<>\"Displacements\"," + "\", Dirn. \" &"+xlApp.range(requestsaddress).offset(0,3).address+ ",\"\")";
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs,4].formula2="=" + xlApp.range(requestsaddress).offset(0,2).address + "& \", \" & "+ xlApp.range(requestsaddress).offset(0,1).address+ " & if("+xlApp.range(requestsaddress).offset(0,1).address +"=\"Actions\"," + "\", Dirn. \" &"+xlApp.range(requestsaddress).offset(0,3).address+ ",\"\")";
       
        chrt.ChartTitle.Text = "='"+ws+"'!" +xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs].address;
        chrt.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatter;
        chrt.Axes(2).MajorGridlines.Delete();
        Microsoft.Office.Interop.Excel.SeriesCollection hlder = chrt.SeriesCollection();

        
        hlder.NewSeries();
        Series nodeseries = chrt.SeriesCollection(1);
        nodeseries.Type = -4169;
        nodeseries.Name = "Nodes";
        nodeseries.MarkerStyle = Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleCircle;
        nodeseries.Format.Line.Visible = 0;
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1].formula2 = "=structuralanalysisnodecoords(" + xlApp.range(cell).offset(2, 0).resize(3,5).address +")";
        nodeseries.XValues = xlApp.workbooks[wb].worksheets[ws].range[xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1].resize[500, 1].address];
        nodeseries.Values = xlApp.workbooks[wb].worksheets[ws].range[xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1, 1].resize[500, 1].address];

        hlder.NewSeries();
        Series elementseries = chrt.SeriesCollection(2);
        elementseries.Type = -4169;
        elementseries.Name = "Elements";
        elementseries.Smooth = false;
        elementseries.MarkerStyle = Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleNone;
        elementseries.Format.Line.Visible = MsoTriState.msoTrue;
        elementseries.Format.Line.Weight=3;
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+4,2].formula2 = "=structuralanalysiselementcoords(" + xlApp.range(cell).offset(2, 0).resize(3,5).address+","+xlApp.range(cell).offset(8, 0).resize(2,10).address +")";
        elementseries.XValues = xlApp.workbooks[wb].worksheets[ws].range[xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+4,2].resize[500, 1].address];
        elementseries.Values = xlApp.workbooks[wb].worksheets[ws].range[xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+4, 3].resize[500, 1].address];

        hlder.NewSeries();
        Series effectseries = chrt.SeriesCollection(3);
        effectseries.Type = -4169;
        effectseries.Name = "='"+ws+"'!" +xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs,4].address;
        effectseries.Smooth = false;
        effectseries.MarkerStyle = Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleNone;
        effectseries.Format.Line.Visible = MsoTriState.msoTrue;;
        effectseries.Format.Line.Weight=2;
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,4].formula2 = "=structuralanalysisgrapheffects(" + xlApp.range(cell).offset(2, 0).resize(3,5).address+","+xlApp.range(cell).offset(8, 0).resize(2,10).address+"," +resultsblockad+"#,"+ xlApp.range(cell).offset(offs-7, 0).resize(1,9).address+ ","+ xlApp.range(cell).offset(13, 0).resize(3,4).address+","+ xlApp.range(cell).offset(19, 0).resize(2,9).address+")";
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,7].resize[1,7].Columns.Autofit();
        effectseries.XValues = xlApp.workbooks[wb].worksheets[ws].range[xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1, 4].resize[2000, 1].address];
        effectseries.Values = xlApp.workbooks[wb].worksheets[ws].range[xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1, 5].resize[2000, 1].address];

        hlder.NewSeries();
        Series rangeseries = chrt.SeriesCollection(4);
        rangeseries.Type = -4169;
        rangeseries.Name = "";
        rangeseries.Smooth = false;
        rangeseries.MarkerStyle = Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleNone;
        rangeseries.Format.Line.Visible = MsoTriState.msoTrue;;
        rangeseries.Format.Line.Weight=0;
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,2].formula2 = "=structuralanalysisgraphrange("+ xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1].address+"#,"+xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,4].address+"#)";
        rangeseries.XValues = xlApp.workbooks[wb].worksheets[ws].range[xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1, 2].resize[3, 1].address];
        rangeseries.Values = xlApp.workbooks[wb].worksheets[ws].range[xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1, 3].resize[3, 1].address];
        chrt.Legend.LegendEntries(4).delete();

        chrt.Axes(1).HasTitle = true;
        chrt.Axes(1).AxisTitle.Text = "x (m)";
        chrt.Axes(2).HasTitle = true;
        chrt.Axes(2).AxisTitle.Text = "y (m)";
        double dummy = chrt.PlotArea.Width; 
        chrt.PlotArea.Width = 0.8*(rght-lft);
        chrt.PlotArea.Height = 0.8*(rght-lft);
      }

      
      //xlApp.workbooks[wb].worksheets[ws].range[cell].offset[20, 10].formula2 = "=SectionProperties(" + xlApp.range(cell).offset(2, 0).resize(2, 15).address + "," + xlApp.range(cell).offset(6, 0).resize(2, 8).address + "," + xlApp.range(cell).offset(10, 0).resize(2, 14).address + "," + xlApp.workbooks[wb].worksheets[ws].range[cell].offset[14].resize[1, 19].address + ")";
    }
    public void createinputblock(dynamic xlApp, string wb, string ws, string cell, int offst, string ttl, string[] hdings, int scndrow, object[] contents1 = null, object[] contents2 = null)
    {
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offst, 0] = ttl;
      for (int i = 0; i < hdings.GetLength(0); i++)
      {
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offst + 1, i] = hdings[i];
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offst + 2, i].style = "Input";
        for (int j=0;j<scndrow;j++)
        {
          xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offst + 3+j, i].style = "Input";
        }
      }
      if (contents1 == null)
      {
        contents1 = new object[] { };
      }
      for (int i = 0; i < contents1.GetLength(0); i++)
      {
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offst + 2, i] = contents1[i];
        if (contents1[i] is bool)
        {
          xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offst + 2, i].validation.add(3, 3, 3, "TRUE,FALSE", 0);
          for (int j=0;j<scndrow;j++)
          {
            xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offst + 3+j, i].validation.add(3, 3, 3, "TRUE,FALSE", 0);
          }
        }
      }
      if (contents2 == null)
      {
        contents2 = new object[] { };
      }
      for (int i = 0; i < contents2.GetLength(0); i++)
      {
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offst + 3, i] = contents2[i];
      }
    }
  }
}

