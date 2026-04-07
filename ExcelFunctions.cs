namespace TESTEXDNA
{
    using ExcelDna.Integration;
    using System;
    using System.Drawing.Drawing2D;
    using System.Windows.Forms.VisualStyles;
    using MathNet.Numerics.LinearAlgebra;

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
            object[,] extracts)

        {
            return controllerclass.controller(nodes,elements,loads,bloads,extracts, new  Dictionary<(int,int),List<double[]>>());
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
            [ExcelArgument(Name = "Extracts")]
            object[,] extracts)

        {
            double[,] tnodes=interfacefunctions.nodefilter(interfacefunctions.filterarrayempties(nodes));
            Dictionary<(int,int),List<double[]>> tnsprings=interfacefunctions.tnsprings(interfacefunctions.filterarrayempties(nodes));
            double[,] telements=interfacefunctions.elementfilter(interfacefunctions.filterarrayempties(elements));
            double[,] tnloads=interfacefunctions.nloadsfilter(interfacefunctions.filterarrayempties(loads),tnodes);
            double[,] tbloads=interfacefunctions.bloadsfilter(interfacefunctions.filterarrayempties(bloads),telements);
            object[,] textracts=interfacefunctions.extractsfilter(interfacefunctions.filterarrayempties(extracts));

            return controllerclass.controller(tnodes,telements,tnloads,tbloads,textracts,tnsprings);
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
            object[,] bloads)

        {
            double[,] tnodes=interfacefunctions.nodefilter(interfacefunctions.filterarrayempties(nodes));
            Dictionary<(int,int),List<double[]>> tnsprings=interfacefunctions.tnsprings(interfacefunctions.filterarrayempties(nodes));
            double[,] telements=interfacefunctions.elementfilter(interfacefunctions.filterarrayempties(elements));
            double[,] tnloads=interfacefunctions.nloadsfilter(interfacefunctions.filterarrayempties(loads),tnodes);
            double[,] tbloads=interfacefunctions.bloadsfilter(interfacefunctions.filterarrayempties(bloads),telements);
            object[,] textracts=new object[,] {{-1,0,0,-1},{-1,1,0,-1},{-1,0,1,-1},{-1,1,1,-1}};

            return controllerclass.controller(tnodes,telements,tnloads,tbloads,textracts,tnsprings);
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
            object[,] requests)

        {
            return graphingtablefunctions.tabularise(results,requests);
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
            object[,] requests)

        {
            return graphingtablefunctions.grapheffects(interfacefunctions.filterarrayempties(nodes),interfacefunctions.filterarrayempties(Elements),results,requests);
        }
        
    }
}
 