using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;

namespace TESTEXDNA
{
  public class MyAddIn : IExcelAddIn
  {
      public void AutoOpen()
      {
          MessageBox.Show("FrameXL is a 2D frame analysis tool built in C# and linked to Excel through 'excel-dna'. For license information, see https://github.com/ShawAAA/FrameXL/. \n\nFrameXL is in early development, and while reasonable efforts have been made to validate results, there may be incorrect or misleading results. As such, refer to the license agreement. Please report any issues to the GitHub repository. \n\nCreated by A. Shaw. \n\nRefer cell notes for guidance on inputting data and interpreting results.");
      }

      public void AutoClose()
      {
      }
  }
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
              <button id='sbutton1' label='New Structure - Pure Text Output' onAction='SetupNewStruct' tag='text'/> 
              <button id='sbutton2' label='New Structure - Excel Graphical/Table Output' onAction='SetupNewStruct' tag='simple'/>
              <button id='sbutton3' label='New Structure - Windowed Graphical Output' onAction='SetupNewStruct' tag='advanced'/>
            </group >
            <group id='group2' label='Structural Tools'>
              <button id='tbutton1' label='New vehicle load' onAction='vehcase'/> 
            </group >
            <group id='group3' label='General Tools'>
              <button id='gbutton1' label='New Animation' onAction='anim'/>
              <button id='gbutton2' label='Run Animations' onAction='animrun'/> 
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
      String cell = xlApp.ActiveCell.offset(1,1).address;
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset(-1,-1).value="FrameXL V0.4. Created by A. Shaw.";
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset(-1,-1).resize(1,2).merge();
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset(-1,1).value="Auto. update matrix calculations:";
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset(-1,1).resize(1,2).merge();
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset(-1,3).style = "Input";
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset(-1,3).validation.add(3, 3, 3, "ON,OFF", 0);
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset(-1,3).value="ON";
      int offs=0;

      string[] nodes1 = new string[] { "x Coord.(m)", "y Coord. (m)", "x fixity","y fixity","Rotation fixity"};
      object[] nodes2 = new object[] { 0, 0, "True","True","True"};
      object[] nodes3 = new object[] { 1, 0, "False","False","False"};
      createinputblock(xlApp, wb, ws, cell, offs, "Node Inputs", nodes1, 2, nodes2,nodes3);
      Comment cmt=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,2].addcomment("Fixities as True/False or with stiffness N/m or Nm/rad. Input spring curves with '-disp/rot,-force/moment;...,...;disp/rot,force/moment'. 'disp/rot' coordinates are in global reference system.");
      cmt.Shape.TextFrame.Characters().Font.Bold = false;
      cmt.Shape.TextFrame.AutoSize = true;
      offs=offs+6;
      
      string[] element1 = new string[] { "Node 1 Index", "Node 2 Index", "EA (N)","EI (Nm2)","End 1 Axial Release","End 1 Shear Release","End 1 Rotational Release","End 2 Axial Release","End 2 Shear Release","End 2 Rotational Release"};
      object[] element2 = new object[] { 1, 2, 1000,1000,"False","False","False","False","False","False"};
      createinputblock(xlApp, wb, ws, cell, offs, "Element Inputs", element1, 1, element2);
      cmt=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,4].addcomment("Releases as True/False or with stiffness N/m or Nm/rad.");
      cmt.Shape.TextFrame.Characters().Font.Bold = false;
      cmt.Shape.TextFrame.AutoSize = true;
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

      string[] combination1 = new string[] { "Combination Load Case Number", "Add Cases / Envelope Cases", "Case Descriptor"};
      object[] combination2 = new object[] { 4, "Add", "1*LC1,2*LC3"};
      createinputblock(xlApp, wb, ws, cell, offs, "Combination Case Inputs", combination1, 1, combination2);
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 2, 1].validation.add(3, 3, 3, "Add,Envelope", 0);
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 3, 1].validation.add(3, 3, 3, "Add,Envelope", 0);
      cmt=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,2].addcomment("Case descriptors are equations of the form '1*LC1,2*LC3' or '1*LC1 to 1*LC3, 3*LC5' where the load case lists are comma delineated and either added or enveloped. Every load case requires a factor. 'to' increment steps also increment the case factor. Can actually be input as '1*1+2*3' but this is bad practice.");
      cmt.Shape.TextFrame.Characters().Font.Bold = false;
      cmt.Shape.TextFrame.AutoSize = true;
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
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[-1,-1].resize[20,10].Columns.Autofit();
      
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs,-1].formula2 = "=if("+  xlApp.workbooks[wb].worksheets[ws].range[cell].offset[-1,3].address+"=\"OFF\",\"Calculations Paused.\",structuralanalysis(" + xlApp.range(cell).offset(2, 0).resize(3,5).address +","+ xlApp.range(cell).offset(8, 0).resize(2,10).address+","+ xlApp.range(cell).offset(13, 0).resize(3,4).address+","+ xlApp.range(cell).offset(19, 0).resize(2,9).address +","+ xlApp.range(cell).offset(24, 0).resize(2,3).address+","+ xlApp.range(cell).offset(offs-3, 0).resize(2,4).address+ "))";
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
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[-1,-1].resize[20,10].Columns.Autofit();
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs,-1].formula2 = "=if("+  xlApp.workbooks[wb].worksheets[ws].range[cell].offset[-1,3].address+"=\"OFF\",\"Calculations Paused.\",structuralanalysissimple(" + xlApp.range(cell).offset(2, 0).resize(3,5).address +","+ xlApp.range(cell).offset(8, 0).resize(2,10).address+","+ xlApp.range(cell).offset(13, 0).resize(3,4).address+","+ xlApp.range(cell).offset(19, 0).resize(2,9).address +","+ xlApp.range(cell).offset(24, 0).resize(2,3).address+ "))";
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs,-1].resize(5,6).HorizontalAlignment=5;
        string resultsblockad=(string)xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs,-1].address();
        offs=offs+5;
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs,5].value="Results Table";
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,5].formula2="=structuralanalysistabularise("+resultsblockad+"#,"+ xlApp.range(cell).offset(offs-7, 0).resize(1,9).address+","+xlApp.range(cell).offset(2, 0).resize(3,5).address+","+xlApp.range(cell).offset(8, 0).resize(2,10).address+","+ xlApp.range(cell).offset(13, 0).resize(3,4).address+","+ xlApp.range(cell).offset(19, 0).resize(2,9).address+")";
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,5].resize[1,6].Columns.Autofit();
        double lft = xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs-5,-1].left;
        double top = xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs-5,-1].top;
        double rght =xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs-5,5].left;
        ChartObject chartob = xlApp.workbooks[wb].worksheets[ws].ChartObjects.Add(lft, top, rght-lft, 0.9*(rght-lft));
        Chart chrt = chartob.Chart;
        chrt.PlotVisibleOnly = false;
        chrt.HasTitle = true;

        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs,-1].formula2="=\"LC\" & "+requestsaddress + "& \", \" & " + xlApp.range(requestsaddress).offset(0,2).address + "& \", \" & "+ xlApp.range(requestsaddress).offset(0,1).address+ " & if("+xlApp.range(requestsaddress).offset(0,1).address +"<>\"Displacements\"," + "\", Dirn. \" &"+xlApp.range(requestsaddress).offset(0,3).address+ ",\"\")";
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs,3].formula2="=" + xlApp.range(requestsaddress).offset(0,2).address + "& \", \" & "+ xlApp.range(requestsaddress).offset(0,1).address+ " & if("+xlApp.range(requestsaddress).offset(0,1).address +"=\"Actions\"," + "\", Dirn. \" &"+xlApp.range(requestsaddress).offset(0,3).address+ ",\"\")";
       
        chrt.ChartTitle.Text = "='"+ws+"'!" +xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs,-1].address;
        chrt.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatter;
        chrt.Axes(2).MajorGridlines.Delete();
        Microsoft.Office.Interop.Excel.SeriesCollection hlder = chrt.SeriesCollection();

        
        hlder.NewSeries();
        Series nodeseries = chrt.SeriesCollection(1);
        nodeseries.Type = -4169;
        nodeseries.Name = "Nodes";
        nodeseries.MarkerStyle = Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleCircle;
        nodeseries.Format.Line.Visible = 0;
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,-1].formula2 = "=structuralanalysisnodecoords(" + xlApp.range(cell).offset(2, 0).resize(3,5).address +")";
        nodeseries.XValues = xlApp.workbooks[wb].worksheets[ws].range[xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,-1].resize[500, 1].address];
        nodeseries.Values = xlApp.workbooks[wb].worksheets[ws].range[xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1, 0].resize[500, 1].address];

        hlder.NewSeries();
        Series elementseries = chrt.SeriesCollection(2);
        elementseries.Type = -4169;
        elementseries.Name = "Elements";
        elementseries.Smooth = false;
        elementseries.MarkerStyle = Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleNone;
        elementseries.Format.Line.Visible = MsoTriState.msoTrue;
        elementseries.Format.Line.Weight=3;
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+4,1].formula2 = "=structuralanalysiselementcoords(" + xlApp.range(cell).offset(2, 0).resize(3,5).address+","+xlApp.range(cell).offset(8, 0).resize(2,10).address +")";
        elementseries.XValues = xlApp.workbooks[wb].worksheets[ws].range[xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+4,1].resize[500, 1].address];
        elementseries.Values = xlApp.workbooks[wb].worksheets[ws].range[xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+4, 2].resize[500, 1].address];

        hlder.NewSeries();
        Series effectseries = chrt.SeriesCollection(3);
        effectseries.Type = -4169;
        effectseries.Name = "='"+ws+"'!" +xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs,3].address;
        effectseries.Smooth = false;
        effectseries.MarkerStyle = Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleNone;
        effectseries.Format.Line.Visible = MsoTriState.msoTrue;;
        effectseries.Format.Line.Weight=2;
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,3].formula2 = "=structuralanalysisgrapheffects(" + xlApp.range(cell).offset(2, 0).resize(3,5).address+","+xlApp.range(cell).offset(8, 0).resize(2,10).address+"," +resultsblockad+"#,"+ xlApp.range(cell).offset(offs-7, 0).resize(1,9).address+ ","+ xlApp.range(cell).offset(13, 0).resize(3,4).address+","+ xlApp.range(cell).offset(19, 0).resize(2,9).address+")";
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,7].resize[1,7].Columns.Autofit();
        effectseries.XValues = xlApp.workbooks[wb].worksheets[ws].range[xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1, 3].resize[2000, 1].address];
        effectseries.Values = xlApp.workbooks[wb].worksheets[ws].range[xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1, 4].resize[2000, 1].address];

        hlder.NewSeries();
        Series rangeseries = chrt.SeriesCollection(4);
        rangeseries.Type = -4169;
        rangeseries.Name = "";
        rangeseries.Smooth = false;
        rangeseries.MarkerStyle = Microsoft.Office.Interop.Excel.XlMarkerStyle.xlMarkerStyleNone;
        rangeseries.Format.Line.Visible = MsoTriState.msoTrue;;
        rangeseries.Format.Line.Weight=0;
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,1].formula2 = "=structuralanalysisgraphrange("+ xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,-1].address+"#,"+xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,3].address+"#)";
        rangeseries.XValues = xlApp.workbooks[wb].worksheets[ws].range[xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1, 1].resize[3, 1].address];
        rangeseries.Values = xlApp.workbooks[wb].worksheets[ws].range[xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1, 2].resize[3, 1].address];
        chrt.Legend.LegendEntries(4).delete();

        chrt.Axes(1).HasTitle = true;
        chrt.Axes(1).AxisTitle.Text = "x (m)";
        chrt.Axes(2).HasTitle = true;
        chrt.Axes(2).AxisTitle.Text = "y (m)";
        double dummy = chrt.PlotArea.Width; 
        chrt.PlotArea.Width = 0.8*(rght-lft);
        chrt.PlotArea.Height = 0.8*(rght-lft);
      }
      else if ((string)tg == "advanced")
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
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[-1,-1].resize[20,10].Columns.Autofit();
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs,-1].formula2 = "=if("+  xlApp.workbooks[wb].worksheets[ws].range[cell].offset[-1,3].address+"=\"OFF\",\"Calculations Paused.\",structuralanalysissimple(" + xlApp.range(cell).offset(2, 0).resize(3,5).address +","+ xlApp.range(cell).offset(8, 0).resize(2,10).address+","+ xlApp.range(cell).offset(13, 0).resize(3,4).address+","+ xlApp.range(cell).offset(19, 0).resize(2,9).address +","+ xlApp.range(cell).offset(24, 0).resize(2,3).address+ "))";
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs,-1].resize(5,6).HorizontalAlignment=5;
        string resultsblockad=(string)xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs,-1].address();
        offs=offs+5;
        Microsoft.Office.Interop.Excel.Range rnge=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs,-1].resize[4,5];
        Microsoft.Office.Interop.Excel.Button btn=xlApp.workbooks[wb].worksheets[ws].Buttons().Add(rnge.Left, rnge.Top, rnge.Width, rnge.Height);
        btn.OnAction = "Showplotmacro";
        btn.Caption = "Render Graph";
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs,-1].formula2="=Structuralanalysisgraphvoid("+resultsblockad+"#,"+xlApp.range(cell).offset(offs-7, 0).resize(1,9).address+","+xlApp.range(cell).offset(2, 0).resize(3,5).address+","+xlApp.range(cell).offset(8, 0).resize(2,10).address+","+ xlApp.range(cell).offset(13, 0).resize(3,4).address+","+ xlApp.range(cell).offset(19, 0).resize(2,9).address+")";
        offs=offs+5;
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs,-1].value="Results Table";
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,-1].formula2="=structuralanalysistabularise("+resultsblockad+"#,"+ xlApp.range(cell).offset(offs-12, 0).resize(1,9).address+","+xlApp.range(cell).offset(2, 0).resize(3,5).address+","+xlApp.range(cell).offset(8, 0).resize(2,10).address+","+ xlApp.range(cell).offset(13, 0).resize(3,4).address+","+ xlApp.range(cell).offset(19, 0).resize(2,9).address+")";
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,-1].resize[1,10].Columns.Autofit();
      }

      
      //xlApp.workbooks[wb].worksheets[ws].range[cell].offset[20, 10].formula2 = "=SectionProperties(" + xlApp.range(cell).offset(2, 0).resize(2, 15).address + "," + xlApp.range(cell).offset(6, 0).resize(2, 8).address + "," + xlApp.range(cell).offset(10, 0).resize(2, 14).address + "," + xlApp.workbooks[wb].worksheets[ws].range[cell].offset[14].resize[1, 19].address + ")";
    }
    public void vehcase(ExcelDna.Integration.CustomUI.IRibbonControl control)
    {
      dynamic xlApp = ExcelDnaUtil.Application;
      String wb = xlApp.ActiveWorkbook.name;
      String ws = xlApp.ActiveSheet.name;
      String cell = xlApp.ActiveCell.address;
      object noderange = xlApp.InputBox("Please select the node range (entire area coloured as input) of all nodes in the structure:","Select Node Input Range",Type:8);
      object elrange = xlApp.InputBox("Please select the element range (entire area coloured as input) of all elements in the structure:","Select Element Input Range",Type:8);
      if (noderange is bool || elrange is bool)
      {
        MessageBox.Show("Selection cancelled.");
        return;
      }
      int offs=0;
      string[] vehinp1 = new string[] { "Element list (elements on chainage)", "Step Increment", "Start node index","Starting load case index"};
      object[] vehinp2 = new object[] { "1to3", 0.5, "1", "1" };
      createinputblock(xlApp, wb, ws, cell, offs, "Vehicle Load Inputs", vehinp1, 0, vehinp2);
      Comment cmt=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,0].addcomment("Element lists can be given as 'All', a list of indexes '1,2,3' or a range '1to3,5to-1'.");
      cmt.Shape.TextFrame.Characters().Font.Bold = false;
      cmt.Shape.TextFrame.AutoSize = true;
      string inprange=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+2,0].address;
      offs=offs+4;

      string[] vehloads1 = new string[] { "Point/Patch","Local/Global/Global Projected","Direction","Axle Point Load Chainage / Chainage of Start of Patch (m)","Point Load Magnitude (N), Start of Patch Magnitude (N/m)","Chainage of End of Patch (m)","End of Patch Magnitude (N/m)"};
      object[] vehloads2 = new object[] { "Point Load","Global Projected","y",0,-1};
      createinputblock(xlApp, wb, ws, cell, offs, "Vehicle Axle Load Inputs", vehloads1, 2, vehloads2);
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 2, 0].validation.add(3, 3, 3, "Point Load,Patch Load", 0);
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 3, 0].validation.add(3, 3, 3, "Point Load,Patch Load", 0);
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 2, 1].validation.add(3, 3, 3, "Local,Global,Global Projected", 0);
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 3, 1].validation.add(3, 3, 3, "Local,Global,Global Projected", 0);
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 2, 2].validation.add(3, 3, 3, "x,y,zz", 0);
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs + 3, 2].validation.add(3, 3, 3, "x,y,zz", 0);
      cmt=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,1].addcomment("Local applied x/axial, y/vertical and zz/rotational. Global Projected Applied across the projected length of the element (N/A for rotation).");
      cmt.Shape.TextFrame.Characters().Font.Bold = false;
      cmt.Shape.TextFrame.AutoSize = true;
      cmt=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,3].addcomment("Chainage should increment in the positive direction, setting in the negative direction places the vehicle loads already on the chainage. Setting chainage 0 sets the first load to be applied at the start node.");
      cmt.Shape.TextFrame.Characters().Font.Bold = false;
      cmt.Shape.TextFrame.AutoSize = true;
      cmt=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,4].addcomment("Magnitude or equation for magnitude. Equations are input as '#index+lc+2*x+y+2', or similar, with properties derived from the position at which the load 'turns-on' evaluated at either the point load or either side of the path.");
      cmt.Shape.TextFrame.Characters().Font.Bold = false;
      cmt.Shape.TextFrame.AutoSize = true;
      string vehrange=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+2,0].resize[3,7].address;
      offs=offs+6;

      string[] elementloads1 = new string[] { "Load Case Number", "Element Applied", "Point/Patch","Local/Global/Global Projected","Direction","Point Load Position / LHS of Patch (%)","Point Load Magnitude (N), LHS Patch Magnitude (N/m)","RHS of Patch (%)","RHS Patch Magnitude (N/m)"};
      object[] elementloads2 = new object[] { };
      createinputblock(xlApp, wb, ws, cell, offs, "Vehicle Load Results", elementloads1, 3, elementloads2,styleselect:"Output");
      //xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+1,0].value="Vehicle Load Calculation";
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+2,0].formula2 = "=structuralanalysisvehicleloads("+((Microsoft.Office.Interop.Excel.Range)noderange).Address+","+((Microsoft.Office.Interop.Excel.Range)elrange).Address+","+inprange+","+vehrange+","+ xlApp.range[inprange].offset[0,1].address+","+xlApp.range[inprange].offset[0,2].address+","+xlApp.range[inprange].offset[0,3].address+")";
      //xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+2,-1].formula2 = "=structuralanalysisarrayindexes("+xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offs+2,0].address+"#)";
    }
    public void anim(ExcelDna.Integration.CustomUI.IRibbonControl control)
    {
      dynamic xlApp = ExcelDnaUtil.Application;
      String wb = xlApp.ActiveWorkbook.name;
      String ws = xlApp.ActiveSheet.name;
      String cell = xlApp.ActiveCell.address;
      object animrange = xlApp.InputBox("Please select the cell value to be updated to run the animation","Select Animation Input Range:",Type:8);
      if (animrange is bool)
      {
        MessageBox.Show("Selection cancelled.");
        return;
      }
      else if (((Microsoft.Office.Interop.Excel.Range)animrange).Cells.CountLarge>1)
      {
        MessageBox.Show("Please select a single cell.");
        return;
      }
      string[] anim1 = new string[] { "Update Cell Range", "Start Value", "End Value","Step Increment","Time delay between steps (ms)","Loop Count (capped at 10)","Wobble"};
      object[] anim2 = new object[] { ((Microsoft.Office.Interop.Excel.Range)animrange).Address, 1, 10,1,100,5,true};
      createinputblock(xlApp, wb, ws, cell, 0, "Animation Inputs", anim1, 0, anim2);
    }
    public void animrun(ExcelDna.Integration.CustomUI.IRibbonControl control)
    {
      dynamic xlApp = ExcelDnaUtil.Application;
      String wb = xlApp.ActiveWorkbook.name;
      String ws = xlApp.ActiveSheet.name;
      object animrange = xlApp.InputBox("Please select theanimation you want to run (select the cell containing the heading 'Animation Inputs')","Select Animation:",Type:8);
      if (animrange is bool)
      {
        MessageBox.Show("Selection cancelled.");
        return;
      }
      else if (((Microsoft.Office.Interop.Excel.Range)animrange).Cells.CountLarge>1)
      {
        MessageBox.Show("Please select a single cell.");
        return;
      }
      string cell=((Microsoft.Office.Interop.Excel.Range)animrange).Address;
      string animcell=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[2,1].value;
      double start=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[2,2].value;
      double end=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[2,3].value;
      double step=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[2,4].value;
      double delay=xlApp.workbooks[wb].worksheets[ws].range[cell].offset[2,5].value;
      string initval=Convert.ToString(xlApp.workbooks[wb].worksheets[ws].range[animcell].value);
      int loop=(int)xlApp.workbooks[wb].worksheets[ws].range[cell].offset[2,6].value;
      bool wobble=(bool)xlApp.workbooks[wb].worksheets[ws].range[cell].offset[2,7].value;
      loop=Math.Max(1,Math.Min(loop,10));
      step=Math.Abs(step)*Math.Sign(end-start);
      while (true)
      {
        for (double val = start; !(val<Math.Min(start, end)||val>Math.Max(start, end)); val += step)
        {
          xlApp.workbooks[wb].worksheets[ws].range[animcell].value = val;
          if (xlApp.ActiveSheet.name == ws)
          {
            foreach (ChartObject chartObj in xlApp.ActiveSheet.ChartObjects())
            {
                chartObj.Chart.Refresh();
                System.Windows.Forms.Application.DoEvents();
            }
          }
          System.Threading.Thread.Sleep((int)delay);
        }
        if (wobble)
        {
          for (double val = end; !(val<Math.Min(start, end)||val>Math.Max(start, end)); val -= step)
          {
            xlApp.workbooks[wb].worksheets[ws].range[animcell].value = val;
            if (xlApp.ActiveSheet.name == ws)
            {
              foreach (ChartObject chartObj in xlApp.ActiveSheet.ChartObjects())
              {
                  chartObj.Chart.Refresh();
                  System.Windows.Forms.Application.DoEvents();
              }
            }
            System.Threading.Thread.Sleep((int)delay);
          }
        }
        loop--;
        if (loop==0)
        {
          break;
        }
      }
      xlApp.workbooks[wb].worksheets[ws].range[animcell]=initval;
    }

    public void createinputblock(dynamic xlApp, string wb, string ws, string cell, int offst, string ttl, string[] hdings, int scndrow, object[] contents1 = null, object[] contents2 = null,string styleselect="Input")
    {
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offst, -1] = ttl;
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offst+1, -1] = "Indexes";
      for (int i = 0; i < hdings.GetLength(0); i++)
      {
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offst + 1, i] = hdings[i];
        xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offst + 2, i].style = styleselect;
        for (int j=0;j<scndrow;j++)
        {
          xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offst + 3+j, i].style = styleselect;
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
      xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offst+2, -1].formula2 = "=structuralanalysisarrayindexes("+xlApp.workbooks[wb].worksheets[ws].range[cell].offset[offst+2].resize[scndrow+1].address+")";
    }
  
  }
  public static class ExcelMacros
  {
      [ExcelCommand(MenuName = "My Tools", MenuText = "Run Macro From Shape")]
      public static void Showplotmacro()
      {
          string cller=ExcelDnaUtil.Application.Caller;
          Microsoft.Office.Interop.Excel.Range rng= ExcelDnaUtil.Application.ActiveSheet.Shapes(cller).TopLeftCell;
          PlotManager.ShowPlot(ExcelDnaUtil.Application.ActiveWorkbook.Name, ExcelDnaUtil.Application.ActiveSheet.Name,rng.Address);
      }
  }

}

