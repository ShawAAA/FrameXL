namespace TESTEXDNA
{
    using ExcelDna.Integration;
    using System.Threading;

    public static class ExcelFunctions
    {
        [ExcelFunction(IsVolatile =false,IsThreadSafe = true,IsHidden =true,Description = "Structural Analysis -unprotected")]
        public static object[,] Structuralanalysisunprotected(
            [ExcelArgument(Name = "Nodes")]
            double[,] nodes,
            [ExcelArgument(Name = "Elements")]
            double[,] elements,
            [ExcelArgument(Name = "Node Loads")]
            double[,] loads,
            [ExcelArgument(Name = "Beam Loads")]
            double[,] bloads,
            [ExcelArgument(Name = "Extracts")]
            object[,] extracts,
            [ExcelArgument(Name = "Load Combinations")]
            object[,] lcomb)

        {
            return controllerclass.controller(nodes,elements,loads,bloads,extracts, new  Dictionary<(int,int),List<double[]>>(),lcomb);
        }

        [ExcelFunction(IsVolatile =false,IsThreadSafe = true,IsHidden =true,Description = "Structural Analysis-text")]
        public static object[,] Structuralanalysis(
            [ExcelArgument(Name = "Nodes")]
            object[,] nodes,
            [ExcelArgument(Name = "Elements")]
            object[,] elements,
            [ExcelArgument(Name = "Node Loads")]
            object[,] loads,
            [ExcelArgument(Name = "Beam Loads")]
            object[,] bloads,
            [ExcelArgument(Name = "Load Combinations")]
            object[,] lcomb,
            [ExcelArgument(Name = "Extracts")]
            object[,] extracts)

        {
            double[,] tnodes=interfacefunctions.nodefilter(interfacefunctions.filterarrayempties(nodes));
            Dictionary<(int,int),List<double[]>> tnsprings=interfacefunctions.tnsprings(interfacefunctions.filterarrayempties(nodes));
            double[,] telements=interfacefunctions.elementfilter(interfacefunctions.filterarrayempties(elements));
            double[,] tnloads=interfacefunctions.nloadsfilter(interfacefunctions.filterarrayempties(loads),tnodes);
            double[,] tbloads=interfacefunctions.bloadsfilter(interfacefunctions.filterarrayempties(bloads),telements,tnodes);
            object[,] textracts=interfacefunctions.extractsfilter(interfacefunctions.filterarrayempties(extracts));
            object[,] tlcomb=interfacefunctions.filterarrayempties(lcomb);

            return controllerclass.controller(tnodes,telements,tnloads,tbloads,textracts,tnsprings,tlcomb);
        }
        [ExcelFunction(IsVolatile =false,IsThreadSafe = true,IsHidden =true,Description = "Structural Analysis-simple")]
        public static object[,] Structuralanalysissimple(
            [ExcelArgument(Name = "Nodes")]
            object[,] nodes,
            [ExcelArgument(Name = "Elements")]
            object[,] elements,
            [ExcelArgument(Name = "Node Loads")]
            object[,] loads,
            [ExcelArgument(Name = "Beam Loads")]
            object[,] bloads,
            [ExcelArgument(Name = "Load Combinations")]
            object[,] lcomb)

        {
            double[,] tnodes=interfacefunctions.nodefilter(interfacefunctions.filterarrayempties(nodes));
            Dictionary<(int,int),List<double[]>> tnsprings=interfacefunctions.tnsprings(interfacefunctions.filterarrayempties(nodes));
            double[,] telements=interfacefunctions.elementfilter(interfacefunctions.filterarrayempties(elements));
            double[,] tnloads=interfacefunctions.nloadsfilter(interfacefunctions.filterarrayempties(loads),tnodes);
            double[,] tbloads=interfacefunctions.bloadsfilter(interfacefunctions.filterarrayempties(bloads),telements,tnodes);
            object[,] textracts=new object[,] {{-1,0,0,-1},{-1,1,0,-1},{-1,0,1,-1},{-1,1,1,-1}};
            object[,] tlcomb=interfacefunctions.filterarrayempties(lcomb);

            return controllerclass.controller(tnodes,telements,tnloads,tbloads,textracts,tnsprings,tlcomb);
        }
        [ExcelFunction(IsVolatile =false,IsThreadSafe = true,IsHidden =true,Description = "Structural Analysis-node graphing")]
        public static object[,] Structuralanalysisnodecoords(
            [ExcelArgument(Name = "Nodes")]
            object[,] nodes)

        {
            return graphingtablefunctions.nodecoords(interfacefunctions.filterarrayempties(nodes));
        }
        [ExcelFunction(IsVolatile =false,IsThreadSafe = true,IsHidden =true,Description = "Structural Analysis-element graphing")]
        public static object[,] Structuralanalysiselementcoords(
            [ExcelArgument(Name = "Nodes")]
            object[,] nodes,
            [ExcelArgument(Name = "Elements")]
            object[,] Elements)

        {
            return graphingtablefunctions.elementcoords(interfacefunctions.filterarrayempties(nodes),interfacefunctions.filterarrayempties(Elements));
        }
        [ExcelFunction(IsVolatile =false,IsThreadSafe = true,IsHidden =true,Description = "Structural Analysis-tabularise")]
        public static object[,] Structuralanalysistabularise(
            [ExcelArgument(Name = "Results")]
            object[,] results,
            [ExcelArgument(Name = "Requested")]
            object[,] requests,
            [ExcelArgument(Name = "Nodes")]
            object[,] nodes,
            [ExcelArgument(Name = "Elements")]
            object[,] elements,
            [ExcelArgument(Name = "Nodes Loads")]
            object[,] loads,
            [ExcelArgument(Name = "Beam Loads")]
            object[,] bloads)

        {
            return graphingtablefunctions.tabularise(results,requests,nodes,elements,loads,bloads);
        }
        [ExcelFunction(IsVolatile =false,IsThreadSafe = true,IsHidden =true,Description = "Structural Analysis-graph effects")]
        public static object[,] Structuralanalysisgrapheffects(
            [ExcelArgument(Name = "Nodes")]
            object[,] nodes,
            [ExcelArgument(Name = "Elements")]
            object[,] Elements,
            [ExcelArgument(Name = "Results")]
            object[,] results,
            [ExcelArgument(Name = "Requested")]
            object[,] requests,
            [ExcelArgument(Name = "Nodes Loads")]
            object[,] loads,
            [ExcelArgument(Name = "Beam Loads")]
            object[,] bloads)

        {
            return graphingtablefunctions.grapheffects(interfacefunctions.filterarrayempties(nodes),interfacefunctions.filterarrayempties(Elements),results,requests,loads,bloads).Item1;
        }
        [ExcelFunction(IsVolatile =false,IsThreadSafe = true,IsHidden =false,Description = "dbg")]
        public static double Structuralanalysisdbg(
            [ExcelArgument(Name = "lc")]
            double[,] lc,
            [ExcelArgument(Name = "index")]
            double[,] index,
            [ExcelArgument(Name = "x")]
            double[,] x,
            [ExcelArgument(Name = "y")]
            double[,] y,
            [ExcelArgument(Name = "zz")]
            double[,] zz,
            [ExcelArgument(Name = "eqns")]
            string[,] eqns)

        {
            return parseclass.parseloadequations(lc[0,0],index[0,0],x[0,0],y[0,0],zz[0,0],eqns[0,0]);
        }
        [ExcelFunction(IsVolatile =false,IsThreadSafe = true,IsHidden =false,Description = "graph range control")]
        public static object[,] Structuralanalysisgraphrange(
            [ExcelArgument(Name = "nodes")]
            object[,] nodes,
            [ExcelArgument(Name = "effects")]
            object[,] effects)

        {
            return graphingtablefunctions.graphrangeset(nodes,effects);
        }
        [ExcelFunction(IsVolatile =false,IsThreadSafe = true,IsHidden =false,Description = "Index Control")]
        public static object[,] Structuralanalysisarrayindexes(
            [ExcelArgument(Name = "array")]
            object[,] inarray)

        {
            return interfacefunctions.arrayindexes(inarray);
        }
        [ExcelFunction(IsVolatile =false,IsThreadSafe = true,IsHidden =false,Description = "Vehicle Loads")]
        public static object[,] Structuralanalysisvehicleloads(
            [ExcelArgument(Name = "nodes")]
            object[,] nodes,
            [ExcelArgument(Name = "elements")]
            object[,] elements,
            [ExcelArgument(Name = "elementlist")]
            string elementlist,
            [ExcelArgument(Name = "vehicle")]
            object[,] vehicle,
            [ExcelArgument(Name = "increment")]
            double increment,
            [ExcelArgument(Name = "startnode")]
            int startnode,
            [ExcelArgument(Name = "initialcase")]
            int initialcase)

        {
            return toolclass.vehicleloads(nodes,elements,elementlist,vehicle,increment,startnode,initialcase);
        }
        [ExcelFunction(IsVolatile =false,IsThreadSafe = false,IsHidden =true,Description = "Structural Analysis-tabularise")]
        public static object Structuralanalysisgraphvoid(
            [ExcelArgument(Name = "Results")]
            object[,] results,
            [ExcelArgument(Name = "Requested")]
            object[,] requests,
            [ExcelArgument(Name = "Nodes")]
            object[,] nodes,
            [ExcelArgument(Name = "Elements")]
            object[,] elements,
            [ExcelArgument(Name = "Nodes Loads")]
            object[,] loads,
            [ExcelArgument(Name = "Beam Loads")]
            object[,] bloads)

        {
            var caller = XlCall.Excel(XlCall.xlfCaller);
            if (caller is ExcelReference reference)
            {
                string rng = "$"+toolclass.ColumnNumberToName(reference.ColumnFirst) + "$"+(reference.RowFirst + 1);
                try
                {
                    PlotManager.takeinp(new List<object[,]>() {results,requests,nodes,elements,loads,bloads});
                    PlotManager.UpdatePlotcontroller();
                }
                catch
                {
                    MessageBox.Show("Graph attempted to update, but has produced an error and closed.");
                    PlotManager.kill();
                }
                
                
            }
            return "RENDER/RECALCULATE";

        }
        
    }
}
 