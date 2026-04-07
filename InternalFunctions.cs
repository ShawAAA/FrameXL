using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using System.Dynamic;
using System.Linq;
using System.Net;
using System.Reflection.PortableExecutable;
using System.Runtime.Intrinsics.X86;
using System.Security.Cryptography;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using ExcelDna.ComInterop;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic.ApplicationServices;
using System.Security.Permissions;
using System.Drawing.Imaging;
using System.Collections;
using MathNet.Numerics.LinearAlgebra.Complex;
using System.Globalization;
using System.Collections.ObjectModel;
using System.Reflection;
using ExcelDna.Integration;
using MathNet.Numerics.Random;
using System.Runtime.InteropServices;

namespace TESTEXDNA
{
    class stiffnessmatcalcs
    {
       
        public static double[,] beamgeom(double[,] tnodes, double[,] telements)
        {
            //length,xvect,yvect,vecangle,EA,EI
            double[,] beamgeom=new double[telements.GetLength(0),6];
            double xvector;
            double yvector;
            double vecmag;
            double vecangle;
            for (int i = 0; i < telements.GetLength(0); i++)
            {
                xvector=tnodes[(int)telements[i,1]-1,0]-tnodes[(int)telements[i,0]-1,0];
                yvector=tnodes[(int)telements[i,1]-1,1]-tnodes[(int)telements[i,0]-1,1];
                vecmag=Math.Sqrt(Math.Pow(xvector,2)+Math.Pow(yvector,2));
                vecangle=Math.Atan2(yvector,xvector);
                beamgeom[i,0]=vecmag;
                beamgeom[i,1]=xvector;
                beamgeom[i,2]=yvector;
                beamgeom[i,3]=vecangle;
                beamgeom[i,4]=telements[i,2];
                beamgeom[i,5]=telements[i,3];
            }
            return beamgeom;
        }
        public static HashSet<int> LClist(double[,] tloads,double[,] tbloads)
        {
            HashSet<int> outp=new HashSet<int>();
            for (int i = 0; i < tloads.GetLength(0); i++)
            {
                outp.Add((int)tloads[i,0]);
            }
            for (int i = 0; i < tbloads.GetLength(0); i++)
            {
                outp.Add((int)tbloads[i,0]);
            }
            return outp;
        }
        public static List<int> springmap(double[,] tnodes)
        {
            List<int> outlist= new List<int>();
            int currow=tnodes.GetLength(0)*3;
            for(int i = 0; i < tnodes.GetLength(0); i++)
            {
                for(int j = 0; j < 3; j++)
                {
                    if (tnodes[i, 2 + j] > 0)
                    {
                        outlist.Add(currow);
                        currow++;
                    }
                    else
                    {
                        outlist.Add(3*i+j);
                    }
                }
            }
            return outlist;
        }
        public static Matrix<double> tmat(double[,] tnodes, double[,] telements,double[,] beamgeom)
        {
            var M=Matrix<double>.Build;
            int originaluscount=tnodes.GetLength(0)*6+telements.GetLength(0)*6;
            int releasecount=0;
            int springcount=0;
            for (int i = 0; i < tnodes.GetLength(0); i++)
            {
                for (int j=0; j < 3; j++)
                {
                    if (tnodes[i, j + 2] > 0)
                    {
                        springcount++;
                    }
                }
            }
            for (int i = 0; i < telements.GetLength(0); i++)
            {
                if (!(telements[i,4]==1 || telements[i,5]==1 || telements[i,6]==1))
                {
                    releasecount=releasecount+3;
                }
                if (!(telements[i,7]==1 || telements[i,8]==1 || telements[i,9]==1))
                {
                    releasecount=releasecount+3;
                }

            }
            int tuscount=originaluscount-releasecount-(tnodes.GetLength(0)*3-springcount);
            Matrix<double> Tarray=M.Dense(originaluscount,tuscount);
            for (int i = 0; i < tnodes.GetLength(0); i++)
            {
                Tarray[3*i,3*i]=1;
                Tarray[3*i+1,3*i+1]=1;
                Tarray[3*i+2,3*i+2]=1;
            }
            int currentcol=tnodes.GetLength(0)*3;
            for (int i = 0; i < tnodes.GetLength(0); i++)
            {
                for (int j = 2; j < 5; j++)
                {
                    if (tnodes[i, j] > 0)
                    {
                        Tarray[tnodes.GetLength(0)*3+i*3+j-2,currentcol]=1;
                        currentcol++;
                    }
                    else
                    {
                       Tarray[tnodes.GetLength(0)*3+i*3+j-2,i*3+j-2]=1;
                    }
                    
                }
                
            }
            int rowcolcontroller;
            int nodenumber;
            double vecangle;
            for (int i = 0; i < telements.GetLength(0); i++)
            {
                vecangle=beamgeom[i,3];
                for (int j = 0; j < 2; j++)
                {
                    rowcolcontroller=(int)Math.Ceiling((double)(j));
                    nodenumber=(int)telements[i,rowcolcontroller];
                    if (telements[i,4+j*3]==1 || telements[i,5+j*3]==1 || telements[i,6+j*3]==1)
                    {
                        switch (telements[i, 4 + j * 3] + telements[i, 5 + j * 3] * 2)
                        {
                            case 0:
                                Tarray[tnodes.GetLength(0)*6+i*6+j*3,(nodenumber-1)*3]=1;
                                Tarray[tnodes.GetLength(0)*6+i*6+j*3+1,(nodenumber-1)*3+1]=1;                            

                                break;
                            case 1:
                                if (vecangle==0 || vecangle == Math.PI)
                                {
                                    Tarray[tnodes.GetLength(0)*6+i*6+j*3,currentcol]=1;
                                    Tarray[tnodes.GetLength(0)*6+i*6+j*3+1,(nodenumber-1)*3+1]=1;                                   

                                }
                                else if(vecangle==Math.PI/2 || vecangle == Math.PI * 1.5)
                                {
                                    Tarray[tnodes.GetLength(0)*6+i*6+j*3,(nodenumber-1)*3]=1;
                                    Tarray[tnodes.GetLength(0)*6+i*6+j*3+1,currentcol+1]=1;
                                }
                                else
                                {

                                    Tarray[tnodes.GetLength(0)*6+i*6+j*3,(nodenumber-1)*3]=1;
                                    Tarray[tnodes.GetLength(0)*6+i*6+j*3,currentcol]=-Math.Sin(vecangle+Math.PI/2);
                                    Tarray[tnodes.GetLength(0)*6+i*6+j*3+1,(nodenumber-1)*3+1]=1;
                                    Tarray[tnodes.GetLength(0)*6+i*6+j*3+1,currentcol]=Math.Cos(vecangle+Math.PI/2);
                                }
                                break;    
                            case 2:
                                if (vecangle==0 || vecangle == Math.PI)
                                {
                                    Tarray[tnodes.GetLength(0)*6+i*6+j*3,(nodenumber-1)*3]=1;
                                    Tarray[tnodes.GetLength(0)*6+i*6+j*3+1,currentcol+1]=1;
                                                                     

                                }
                                else if(vecangle==Math.PI/2 || vecangle == Math.PI * 1.5)
                                {
                                    Tarray[tnodes.GetLength(0)*6+i*6+j*3,currentcol]=1;
                                    Tarray[tnodes.GetLength(0)*6+i*6+j*3+1,(nodenumber-1)*3+1]=1;  
                                }
                                else
                                {
                                    Tarray[tnodes.GetLength(0)*6+i*6+j*3,(nodenumber-1)*3]=1;
                                    Tarray[tnodes.GetLength(0)*6+i*6+j*3,currentcol]=-Math.Sin(vecangle);
                                    Tarray[tnodes.GetLength(0)*6+i*6+j*3+1,(nodenumber-1)*3+1]=1;
                                    Tarray[tnodes.GetLength(0)*6+i*6+j*3+1,currentcol]=Math.Cos(vecangle);
                                }
                                break;
                            case 3:
                                Tarray[tnodes.GetLength(0)*6+i*6+j*3,currentcol]=1;
                                Tarray[tnodes.GetLength(0)*6+i*6+j*3+1,currentcol+1]=1;
                                break;                                    
                        }
                        if (telements[i, 6 + j * 3] == 0)
                        {
                            Tarray[tnodes.GetLength(0)*6+i*6+j*3+2,(nodenumber-1)*3+2]=1;
                        }
                        else
                        {
                            Tarray[tnodes.GetLength(0)*6+i*6+j*3+2,currentcol+2]=1;
                        }
                        currentcol=currentcol+3;
                    }
                    else
                    {
                        for(int x = 0; x < 3; x++)
                        {
                            Tarray[tnodes.GetLength(0)*6+i*6+j*3+x,(nodenumber-1)*3+x]=1;
                        }
                    }
            
                    
                }
                
                
            }
            Vector<double> colsum=Tarray.ColumnSums();
            for (int i = Tarray.ColumnCount - 1; i >= 0; i--)
            {
                if (colsum[i]==0)
                {
                    Tarray=Tarray.RemoveColumn(i);
                }
            }
            return Tarray;
        }
        public static Matrix<double> kmat(double[,] tnodes, double[,] telements,double[,] beamgeom)
        {
            var M=Matrix<double>.Build;
            Matrix<double> kmat=M.Dense(tnodes.GetLength(0)*6+telements.GetLength(0)*6,tnodes.GetLength(0)*6+telements.GetLength(0)*6);
            for (int i = 0; i < tnodes.GetLength(0); i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    if (tnodes[i, 2 + j] > 0)
                    {
                        kmat[i*3+j,i*3+j]=tnodes[i, 2 + j];
                        kmat[i*3+j,tnodes.GetLength(0)*3+i*3+j]=-tnodes[i, 2 + j];
                        kmat[tnodes.GetLength(0)*3+i*3+j,i*3+j]=-tnodes[i, 2 + j];
                        kmat[tnodes.GetLength(0)*3+i*3+j,tnodes.GetLength(0)*3+i*3+j]=tnodes[i, 2 + j];
                    }
                }
                
            }
            Matrix<double> beamlocal;
            for (int i = 0; i < telements.GetLength(0); i++)
            {
                beamlocal=stiffnessmatcalcs.beamstiff(beamgeom[i,0],beamgeom[i,3],telements[i,2],telements[i,3]);
                for (int j = 0; j < 6; j++)
                {
                    for (int k = 0; k < 6; k++)
                    {
                        kmat[tnodes.GetLength(0)*6+i*6+j,tnodes.GetLength(0)*6+i*6+k]=beamlocal[j,k];
                    }
                
                }
            }
            return kmat;
        }
        public static Matrix<double> beamlocal(double EA,double EI,double L)
        {
            var M=Matrix<double>.Build;
            double Eaonl=EA/L;
            double Eionl3=12*EI/Math.Pow(L,3);
            double Eionl2=6*EI/Math.Pow(L,2);
            double Eionl=2*EI/L;
            double[,] ret={{Eaonl,0,0,-Eaonl,0,0},
                           {0,Eionl3,Eionl2,0,-Eionl3,Eionl2},
                           {0,Eionl2,2*Eionl,0,-Eionl2,Eionl},
                           {-Eaonl,0,0,Eaonl,0,0},
                           {0,-Eionl3,-Eionl2,0,Eionl3,-Eionl2},
                           {0,Eionl2,Eionl,0,-Eionl2,2*Eionl}};
            return M.DenseOfArray(ret);
        }
        public static Matrix<double> beamglobaltransform(double alpha)
        {
            var M=Matrix<double>.Build;
            double c=Math.Cos(alpha);
            double s=Math.Sin(alpha);
            double[,] ret={{c,s,0,0,0,0},
                           {-s,c,0,0,0,0},
                           {0,0,1,0,0,0},
                           {0,0,0,c,s,0},
                           {0,0,0,-s,c,0},
                           {0,0,0,0,0,1}};
            return M.DenseOfArray(ret);
        }
        public static Matrix<double> beamstiff(double L,double alpha,double EA,double EI)
        {
            Matrix<double> beamlocal=stiffnessmatcalcs.beamlocal(EA,EI,L);
            Matrix<double> beamtransform=stiffnessmatcalcs.beamglobaltransform(alpha);
            Matrix<double> bgmatrix=((beamtransform.Transpose()).Multiply(beamlocal)).Multiply(beamtransform);
            return bgmatrix;
        }
        public static double[] filtermap(double[,] tnodes,int springcount)
        {
            double[] filter=new double[tnodes.GetLength(0)*3+springcount];
            for (int i = 0; i < tnodes.GetLength(0); i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    if (tnodes[i, j + 2] != -1)
                    {
                        filter[i*3+j]=1;
                    }
                    else
                    {
                        filter[i*3+j]=0;
                    }
                }
            }

            return filter;
        }
        public static Matrix<double> kmatreduce(Matrix<double> kmat,double[] filtermap)
        {
            Matrix<double> kmatout=kmat;
            //for (int i = kmat.RowCount-1; i >= tnodes.GetLength(0) * 3; i--)
            //{
                //kmatout=kmatout.RemoveRow(i);
                //kmatout=kmatout.RemoveColumn(i);
            //} 
           for (int i = filtermap.GetLength(0)-1; i >= 0; i--)
            {
                if (filtermap[i] == 0)
                {
                    kmatout=kmatout.RemoveRow(i);
                    kmatout=kmatout.RemoveColumn(i);
                }
            }
            return kmatout;
        }
        public static Vector<double> Fvector(double[,]tloads,int totaldof,Dictionary<int,List<int>> nodeloadindexes,double[,] beamgeom,double[,] telements,Dictionary<int,List<double[,]>> beamloadsholderindex)
        {
            var v=Vector<double>.Build;
            Vector<double> Fvector=v.Dense(totaldof);
            int i;
            foreach (int key in nodeloadindexes.Keys)
            {
                for(int j=0; j < nodeloadindexes[key].Count; j++)
                {
                    i=nodeloadindexes[key][j];
                    Fvector[(int)((tloads[i,1]-1)*3+tloads[i,2])]=Fvector[(int)((tloads[i,1]-1)*3+tloads[i,2])]+tloads[i,3];
                }
                
            }
            double[,] beamloads;
            double s;
            double c;
            int elstart=totaldof-telements.GetLength(0)*6;
            foreach (int key in beamloadsholderindex.Keys)
            {
                s=Math.Sin(beamgeom[key-1,3]);
                c=Math.Cos(beamgeom[key-1,3]);
                for(int j=0; j < beamloadsholderindex[key].Count; j++)
                {
                    beamloads=beamloadsholderindex[key][j];
                    for(int k=0; k < beamloads.GetLength(0); k++)
                    {
                        Fvector[elstart+(key-1)*6]=Fvector[elstart+(key-1)*6]+beamloads[k,7]*c-beamloads[k,8]*s;
                        Fvector[elstart+(key-1)*6+1]=Fvector[elstart+(key-1)*6+1]+beamloads[k,7]*s+beamloads[k,8]*c;
                        Fvector[elstart+(key-1)*6+2]=Fvector[elstart+(key-1)*6+2]+beamloads[k,9];
                        Fvector[elstart+(key-1)*6+3]=Fvector[elstart+(key-1)*6+3]+beamloads[k,10]*c-beamloads[k,11]*s;
                        Fvector[elstart+(key-1)*6+4]=Fvector[elstart+(key-1)*6+4]+beamloads[k,10]*s+beamloads[k,11]*c;
                        Fvector[elstart+(key-1)*6+5]=Fvector[elstart+(key-1)*6+5]+beamloads[k,12];
                    }
                }
                
            }
            return Fvector;
        }
        public static int springcount(double[,] tnodes)
        {
            int springcount=0;
            for (int i = 0; i < tnodes.GetLength(0); i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    if (tnodes[i, j + 2] > 0)
                    {
                        springcount++;
                    }
                }
            }
            return springcount;
        }
        public static Vector<double> filterby(Vector<double> vect,double[] filter)
        {
            
            Matrix<double> colholder=vect.ToColumnMatrix();
            for (int i = filter.GetLength(0) - 1; i >= 0;i--)
            {
                if (filter[i] == 0)
                {
                    colholder=colholder.RemoveRow(i);
                }
            }
            return colholder.Column(0);
        }
        public static Vector<double> unfilterby(Vector<double> vect,double[] filter,int totdof)
        {
            var v=Vector<double>.Build;
            int currentofvec=0;
            Vector<double> outvec=v.Dense(totdof);
            for (int i = 0; i <filter.GetLength(0);i++)
            {
                if (filter[i] == 1)
                {
                    outvec[i]=vect[currentofvec];
                    currentofvec++;
                }
                else
                {
                    outvec[i]=0;
                }
            }
            for (int i = currentofvec; i < vect.Count; i++)
            {
                outvec[filter.GetLength(0)+i-currentofvec]=vect[i];
            }
            return outvec;
        }
        public static Dictionary<int,Dictionary<int,List<int>>> nodeloadsholder(double[,] tloads)
        {
            Dictionary<int,Dictionary<int,List<int>>> loadholder = new Dictionary<int,Dictionary<int,List<int>>>();
            for (int i = 0; i < tloads.GetLength(0); i++)
            {
                if (loadholder.ContainsKey((int)tloads[i, 0]))
                {
                    if (loadholder[(int)tloads[i, 0]].ContainsKey((int)tloads[i, 1]))
                    {
                        loadholder[(int)tloads[i, 0]][(int)tloads[i, 1]].Add(i);
                    }
                    else
                    {
                        loadholder[(int)tloads[i, 0]].Add((int)tloads[i, 1],new List<int> {i});
                    }

                }
                else
                {
                    loadholder.Add((int)tloads[i, 0],new Dictionary<int,List<int>> {{(int)tloads[i, 1],new List<int> {i}}});
                }
            }
            return loadholder;
        }
        public static Dictionary<int,Dictionary<int,List<double[,]>>> beamloadsholder(double[,] tbloads,double[,] beamgeom)
        {
            var M=Matrix<double>.Build;
            Dictionary<int,Dictionary<int,List<double[,]>>> loadholder = new Dictionary<int,Dictionary<int,List<double[,]>>>();
            Matrix<double> tbloadsmat=M.DenseOfArray(tbloads);
            Matrix<double> beamgeommat=M.DenseOfArray(beamgeom);
            double[,] beamarray;
            for (int i = 0; i < tbloads.GetLength(0); i++)
            {
                beamarray=stiffnessmatcalcs.beamloadconverted(tbloadsmat.Row(i).ToArray(),beamgeommat.Row((int)tbloads[i, 1]-1).ToArray());
                if (loadholder.ContainsKey((int)tbloads[i, 0]))
                {
                    if (loadholder[(int)tbloads[i, 0]].ContainsKey((int)tbloads[i, 1]))
                    {
                        loadholder[(int)tbloads[i, 0]][(int)tbloads[i, 1]].Add(beamarray);
                    }
                    else
                    {
                        loadholder[(int)tbloads[i, 0]].Add((int)tbloads[i, 1],new List<double[,]> {beamarray});
                    }

                }
                else
                {
                    loadholder.Add((int)tbloads[i, 0],new Dictionary<int,List<double[,]>> {{(int)tbloads[i, 1],new List<double[,]> {beamarray}}});
                }
            }
            return loadholder;
        }
        public static double[,] beamloadconverted(double[] bloadsrow,double[] beamgeomrow)
        {
            double projectionscalar=1;
            double axfraction;
            double bendfraction;
            double[,] returnholder;
            double[,] outholder;
            double scle;
            if (bloadsrow[4] == 2)
            {
                //return rotational double[]
                return fixedendcreator.bendingmomentloads(bloadsrow,beamgeomrow);
            }
            if (bloadsrow[3] == 2)
            {
                
                if (bloadsrow[4] == 0)
                {
                    projectionscalar=Math.Abs(beamgeomrow[2]/beamgeomrow[0]);
                }
                else
                {
                    projectionscalar=Math.Abs(beamgeomrow[1]/beamgeomrow[0]);
                }
            }
            if (bloadsrow[3]==0 || beamgeomrow[3] % (Math.PI / 2) == 0)
            {
                int effectivedirn;
                if (bloadsrow[3] == 0)
                {
                    effectivedirn=(int)bloadsrow[4];
                    scle=1;
                }
                else
                {
                    effectivedirn=(int)((Math.Abs(2*((beamgeomrow[3]/Math.PI) % 1))+bloadsrow[4]) % 2);
                    if (bloadsrow[4] == 0)
                    {
                        if (beamgeomrow[3]==0 || beamgeomrow[3]==-Math.PI/2 || beamgeomrow[3] == 3 * Math.PI / 2)
                        {
                            scle=1;
                        }
                        else
                        {
                            scle=-1;
                        }
                    }
                    else
                    {
                        if (beamgeomrow[3]==0 || beamgeomrow[3]==Math.PI/2)
                        {
                            scle=1;
                        }
                        else
                        {
                            scle=-1;
                        }
                    }
                    
                }
                //return aligned double[]
                if (effectivedirn == 0)
                {
                    return fixedendcreator.axialloads(bloadsrow,beamgeomrow,projectionscalar,scle);
                }
                else
                {
                    return fixedendcreator.bendingloads(bloadsrow,beamgeomrow,projectionscalar,scle);
                }

            }
            else
            {
                if (bloadsrow[4] == 0)
                {
                    axfraction=beamgeomrow[1]/beamgeomrow[0];
                    bendfraction=-beamgeomrow[2]/beamgeomrow[0];
                }
                else
                {
                    axfraction=beamgeomrow[2]/beamgeomrow[0];
                    bendfraction=beamgeomrow[1]/beamgeomrow[0];
                }
                //return two-direction double[,]
                outholder= new double[2,13];
                returnholder= fixedendcreator.axialloads(bloadsrow,beamgeomrow,projectionscalar,axfraction);
                for (int i = 0; i < outholder.GetLength(1); i++)
                {
                    outholder[0,i]=returnholder[0,i];
                }
                returnholder= fixedendcreator.bendingloads(bloadsrow,beamgeomrow,projectionscalar,bendfraction);
                for (int i = 0; i < outholder.GetLength(1); i++)
                {
                    outholder[1,i]=returnholder[0,i];
                }
                return outholder;
            }
        }
        public static Dictionary<int,Dictionary<int,SortedSet<double>>> beamchapoints(Dictionary<int,Dictionary<int,List<double[,]>>> beamloadsholder)
        {
            Dictionary<int,Dictionary<int,SortedSet<double>>> chapoint=new Dictionary<int,Dictionary<int,SortedSet<double>>>();
            SortedSet<double> tempchaout;
            SortedSet<double> tempchain;
            Dictionary<int,SortedSet<double>> tempbeamdict;
            foreach (int lc in beamloadsholder.Keys)
            {
                tempbeamdict=new Dictionary<int, SortedSet<double>>();
                foreach (int beam in beamloadsholder[lc].Keys)
                {
                    tempchain= new SortedSet<double>
                    {
                        0,
                        1
                    };
                    for (int loadindex=0;loadindex< beamloadsholder[lc][beam].Count;loadindex++)
                    {

                        for (int rw = 0; rw < beamloadsholder[lc][beam][loadindex].GetLength(0); rw++)
                        {
                           
                            tempchaout=stiffnessmatcalcs.loadchapoints(beamloadsholder[lc][beam][loadindex],rw);
                            tempchain.UnionWith(tempchaout);
                        }
                        
                    }
                    SortedSet<double> filterchapoints=new SortedSet<double>();
                    filterchapoints.Add(tempchain.ElementAt(0));
                    for (int i = 1; i < tempchain.Count - 1; i++)
                    {
                        if(!(Math.Abs(tempchain.ElementAt(i)-tempchain.ElementAt(i-1))<0.000001 && Math.Abs(tempchain.ElementAt(i+1)-tempchain.ElementAt(i))<0.000001))
                        {
                            filterchapoints.Add(tempchain.ElementAt(i));
                        }
                    }
                    filterchapoints.Add(tempchain.ElementAt(tempchain.Count-1));
                    tempbeamdict.Add(beam,filterchapoints);
                }
                chapoint.Add(lc,tempbeamdict);
            }
            
            return chapoint;
        }
        public static SortedSet<double> loadchapoints(double[,] beamloads, int rw)
        {
            SortedSet<double> chapoints= new SortedSet<double>();
            if (beamloads[rw, 0] == 0)
            {
                if (beamloads[rw, 3] != 0)
                {
                    chapoints.Add(beamloads[rw,3]-0.0000001);
                }
                if (beamloads[rw, 3] != 1)
                {
                    chapoints.Add(beamloads[rw,3]+0.0000001);
                }
            }
            else
            {
                double a=beamloads[rw,3];
                double b=beamloads[rw,5];
                double x=Math.Max(1,Math.Ceiling(Math.Log2(Math.Abs((b-a))/0.2)));
                double lim=Math.Pow(2,x);
                for(int i = 0; i <= lim; i++)
                {
                    chapoints.Add(a+(b-a)*i/lim);
                }
            }
            
            return chapoints;
        }
        public static List<int> lcmemberstringread(string txt,int cap,List<int> kys)
        {
            txt=Regex.Replace(txt,@"\s+",string.Empty);
            int startstring=0;
            int nextindex;
            List<int> outlist= new List<int>();
            List<int> holder;
            if (txt == "-1")
            {
                foreach (int ky in kys)
                {
                    outlist.Add(ky);
                }
            }
            else
            {
                nextindex=txt.IndexOf(",", startstring);
                while (nextindex != -1)
                {
                    holder=stiffnessmatcalcs.lcmemberfragment(txt.Substring(startstring,nextindex-startstring),cap);
                    for (int i = 0; i < holder.Count; i++)
                    {
                        outlist.Add(holder[i]);
                    }
                    startstring=nextindex+1;
                    nextindex=txt.IndexOf(",", startstring);
                }
                holder=stiffnessmatcalcs.lcmemberfragment(txt.Substring(startstring),cap);
                for (int i = 0; i < holder.Count; i++)
                {
                    outlist.Add(holder[i]);
                }
            }

            return outlist;
        }
        public static List<int> lcmemberfragment(string txt,int cap)
        {
            txt=txt.ToLower();
            List<int> outlist= new List<int>();
            int index=txt.IndexOf("to");
            if (index != -1)
            {
                int ind1=int.Parse(txt.Substring(0,index));
                if (ind1 == -1)
                {
                    ind1=cap;
                }
                int ind2=int.Parse(txt.Substring(index+2));
                if (ind2 == -1)
                {
                    ind2=cap;
                }
                int dirn;
                if (ind2 > ind1)
                {
                    dirn=1;
                }
                else
                {
                    dirn=-1;
                }
                for (int i = ind1; dirn*i <= dirn*ind2; i = i + dirn)
                {
                    outlist.Add(i);
                }
            }
            else
            {
                outlist.Add(int.Parse(txt.Substring(0,txt.Count())));
            }
            
            return outlist;
        }
        public static string[,] extraction(object[,] extracts,Dictionary<int,Matrix<double>> resultsholder,Dictionary<int,Matrix<double>> resultsholder2,Dictionary<int,Dictionary<int,List<int>>> nodeloadsholder,Dictionary<int,List<Vector<double>>> fextholder,double[,] tloads, int nodecount,int elementcount,List<int> springmap,double[,] beamgeom,Dictionary<int,Dictionary<int,List<double[,]>>> beamloadsholder, Dictionary<int,Dictionary<int,SortedSet<double>>> beamchapoints)
        {
            string[,] outarray=new string[extracts.GetLength(0),1];
            List<int> lclist;
            List<int> memberlist;
            int indexcount;
            string holder="";
            List<int> inlist;
            for (int i=0; i < extracts.GetLength(0); i++)
            {
                lclist = stiffnessmatcalcs.lcmemberstringread(extracts[i,0].ToString()?? string.Empty, resultsholder.Keys.ToList().Max(),resultsholder.Keys.ToList());
                if (Convert.ToInt32(extracts[i, 2]) == 0)
                {
                    indexcount=nodecount;
                }
                else
                {
                    indexcount=elementcount;
                }
                inlist=new List<int>();
                for (int j = 1; j <= indexcount; j++)
                {
                    inlist.Add(j);
                }
                memberlist = stiffnessmatcalcs.lcmemberstringread(extracts[i,3].ToString()?? string.Empty, indexcount,inlist);

                switch (Convert.ToInt32(extracts[i,1]) + Convert.ToInt32(extracts[i,2]) * 2)
                {
                    case 0:
                        holder="Node Reactions$"+stiffnessmatcalcs.actionnodes(resultsholder2,fextholder,lclist,memberlist,springmap);                 

                        break;
                    case 1:
                        holder="Node Displacements$"+stiffnessmatcalcs.dispnodes(resultsholder,lclist,memberlist);

                        break;
                    case 2:
                        holder="Element Actions$"+stiffnessmatcalcs.actionelement(beamgeom,resultsholder,beamloadsholder,beamchapoints,lclist,memberlist,nodecount*6+elementcount*6);

                        break;
                    case 3:
                        holder="Element Displacements$"+stiffnessmatcalcs.dispelement(beamgeom,resultsholder,beamloadsholder,beamchapoints,lclist,memberlist,nodecount*6+elementcount*6);

                        break;
                }
                if (holder.IndexOf("NaN") != -1)
                {
                    return new string[,] {{"Calculations are invalid and have produced one or more results of 'NaN'/infinite, confirm that all degrees of freedom have been adequately constrained."}};
                }
                outarray[i,0]=holder;
            }
            double ln=0;
            double maxlength=32767;
            for (int i = 0; i < outarray.GetLength(0); i++)
            {
                ln=Math.Max(ln,outarray[i,0].Length);
            }
            int cols=(int)Math.Ceiling(ln/maxlength);
            string[,] outarray2=new string[outarray.GetLength(0),cols];
            int strlength;
            for (int i = 0; i < outarray.GetLength(0); i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    strlength=(int)Math.Max(0,outarray[i,0].Length-(j*maxlength));
                    if (strlength > 0)
                    {
                        outarray2[i,j]=outarray[i,0].Substring((int)(j*maxlength),(int)Math.Min(maxlength,strlength));
                    }
                    else
                    {
                        outarray2[i,j]="";
                    }
                    
                }
            }

            return outarray2;
        }
        public static string dispnodes(Dictionary<int,Matrix<double>> resultsholder,List<int> lclist, List<int> nodelist)
        {
            string outstring="";
            string tempstring;
            double moveholder;
            Matrix<double> lcunpacked;
            for (int i = 0; i < lclist.Count; i++)
            {
                lcunpacked=resultsholder[lclist[i]];
                for (int j = 0; j < nodelist.Count; j++)
                {
                    tempstring=lclist[i]+"~"+nodelist[j]+"/";
                    for (int k = 0; k < 3; k++)
                    {
                        moveholder=lcunpacked[(nodelist[j]-1)*3+k,0];
                        if (Math.Abs(moveholder) < Math.Pow(10, -12))
                        {
                            moveholder=0;
                        }
                        tempstring=tempstring+string.Format("{0:G4}",moveholder)+",";
                    }
                    outstring+=tempstring;
                    outstring=outstring.Substring(0,outstring.Length-1)+";";
                }
                outstring=outstring.Substring(0,outstring.Length-1)+":";
            }
            outstring=outstring.Substring(0,outstring.Length-1);

            return outstring;
        }
        public static string actionnodes(Dictionary<int,Matrix<double>> resultsholder2,Dictionary<int,List<Vector<double>>> fextholder,List<int> lclist, List<int> nodelist,List<int> springmap)
        {
            string outstring="";
            string tempstring;
            double reacholder;
            int currnode;
            int mappedindex;
            for (int i = 0; i < lclist.Count; i++)
            {
                for (int j = 0; j < nodelist.Count; j++)
                {
                    tempstring=lclist[i]+"~"+nodelist[j]+"/";
                    currnode=nodelist[j];
                    for (int k = 0; k < 3; k++)
                    {
                        mappedindex=springmap[(currnode-1)*3+k];
                        reacholder=-fextholder[lclist[i]][1][mappedindex]+resultsholder2[lclist[i]][mappedindex,1];
                        if (Math.Abs(reacholder) < Math.Pow(10, -12))
                        {
                            reacholder=0;
                        }
                        tempstring=tempstring+string.Format("{0:G4}",reacholder)+",";
                    }
                    outstring+=tempstring;
                    outstring=outstring.Substring(0,outstring.Length-1)+";";
                }
                outstring=outstring.Substring(0,outstring.Length-1)+":";
            }
            outstring=outstring.Substring(0,outstring.Length-1);
            
            return outstring;
        }
        public static string dispelement(double[,] beamgeom,Dictionary<int,Matrix<double>> resultsholder,Dictionary<int,Dictionary<int,List<double[,]>>> beamloadsholder, Dictionary<int,Dictionary<int,SortedSet<double>>> beamchapoints,List<int> lclist, List<int> elementlist, int totaldof)
        {
            string outstring="";
            string tempstring;
            Matrix<double> effectmatrix;
            double L;
            double s;
            double c;
            List<double[,]> bloads;
            SortedSet<double> cha;
            double axialforce;
            double axialtrans;
            double bendingtransforce;
            double bendingmoment;
            double bendingtrans;
            double rotation;
            int beamcount=beamgeom.GetLength(0);
            int beamzeroindex=totaldof-beamcount*6;
            bool isvalid;
            double moveholder;
            int n=4;
            for (int i = 0; i < lclist.Count; i++)
            {
                for (int j = 0; j < elementlist.Count; j++)
                {
                    tempstring=lclist[i]+"~"+elementlist[j]+"/";
                    L=beamgeom[elementlist[j]-1,0];
                    s=Math.Sin(beamgeom[elementlist[j]-1,3]);
                    c=Math.Cos(beamgeom[elementlist[j]-1,3]);
                    axialforce=-(resultsholder[lclist[i]][beamzeroindex+(elementlist[j]-1)*6,1]*c+resultsholder[lclist[i]][beamzeroindex+(elementlist[j]-1)*6+1,1]*s);
                    axialtrans=(resultsholder[lclist[i]][beamzeroindex+(elementlist[j]-1)*6,0]*c+resultsholder[lclist[i]][beamzeroindex+(elementlist[j]-1)*6+1,0]*s);
                    bendingtransforce=-resultsholder[lclist[i]][beamzeroindex+(elementlist[j]-1)*6,1]*s+resultsholder[lclist[i]][beamzeroindex+(elementlist[j]-1)*6+1,1]*c;
                    bendingmoment=-resultsholder[lclist[i]][beamzeroindex+(elementlist[j]-1)*6+2,1];
                    bendingtrans=-resultsholder[lclist[i]][beamzeroindex+(elementlist[j]-1)*6,0]*s+resultsholder[lclist[i]][beamzeroindex+(elementlist[j]-1)*6+1,0]*c;
                    rotation=resultsholder[lclist[i]][beamzeroindex+(elementlist[j]-1)*6+2,0];
                    isvalid=false;
                    if (beamloadsholder.ContainsKey(lclist[i]))
                    {
                        if (beamloadsholder[lclist[i]].ContainsKey(elementlist[j]))
                        {
                            isvalid=true;
                        }
                    }
                    cha=new SortedSet<double>();
                    for (int k = 0; k <= Math.Pow(2, n); k++)
                    {
                        cha.Add(k/Math.Pow(2,n));
                    }
                    if (isvalid)
                    {
                        cha.UnionWith(beamchapoints[lclist[i]][elementlist[j]]);
                        for (int k = cha.Count-1; k >= 1; k--)
                        {
                            if (Math.Abs(cha.ElementAt(k) - cha.ElementAt(k - 1)) < 0.0001)
                            {
                                cha.Remove(cha.ElementAt(k));
                            }
                        }
                        bloads=beamloadsholder[lclist[i]][elementlist[j]];
                    }
                    else
                    {
                        bloads=new List<double[,]>();
                    }
                    effectmatrix=stiffnessmatcalcs.beamdispcalc(bloads,cha,axialforce,axialtrans,bendingtransforce,bendingmoment,rotation,bendingtrans,beamgeom[elementlist[j]-1,4],beamgeom[elementlist[j]-1,5],beamgeom[elementlist[j]-1,0]);
                    for (int k = 0; k < effectmatrix.RowCount; k++)
                    {
                        tempstring=tempstring+string.Format("{0:G4}", cha.ElementAt(k))+"|"+string.Format("{0:G4}", cha.ElementAt(k)*L)+"#";
                        for (int l = 0; l < 3; l++)
                        {
                            moveholder=effectmatrix[k,l];
                            if (Math.Abs(moveholder) < Math.Pow(10, -12))
                            {
                                moveholder=0;
                            }
                            tempstring=tempstring+string.Format("{0:G4}",moveholder)+",";
                        }
                        tempstring=tempstring.Substring(0,tempstring.Length-1)+"+";
                    }
                    outstring+=tempstring;
                    outstring=outstring.Substring(0,outstring.Length-1)+";";
                }
                outstring=outstring.Substring(0,outstring.Length-1)+":";
            }
            outstring=outstring.Substring(0,outstring.Length-1);
            
            return outstring;
        }
        public static string actionelement(double[,] beamgeom,Dictionary<int,Matrix<double>> resultsholder,Dictionary<int,Dictionary<int,List<double[,]>>> beamloadsholder, Dictionary<int,Dictionary<int,SortedSet<double>>> beamchapoints,List<int> lclist, List<int> elementlist, int totaldof)
        {
            string outstring="";
            string tempstring;
            Matrix<double> effectmatrix;
            double L;
            double s;
            double c;
            List<double[,]> bloads;
            SortedSet<double> cha;
            double axial;
            double bending;
            double rotation;
            int beamcount=beamgeom.GetLength(0);
            int beamzeroindex=totaldof-beamcount*6;
            bool isvalid;
            double moveholder;
            double tempcha;
            for (int i = 0; i < lclist.Count; i++)
            {
                for (int j = 0; j < elementlist.Count; j++)
                {
                    tempstring=lclist[i]+"~"+elementlist[j]+"/";
                    L=beamgeom[elementlist[j]-1,0];
                    s=Math.Sin(beamgeom[elementlist[j]-1,3]);
                    c=Math.Cos(beamgeom[elementlist[j]-1,3]);
                    axial=resultsholder[lclist[i]][beamzeroindex+(elementlist[j]-1)*6,1]*c+resultsholder[lclist[i]][beamzeroindex+(elementlist[j]-1)*6+1,1]*s;
                    bending=-resultsholder[lclist[i]][beamzeroindex+(elementlist[j]-1)*6,1]*s+resultsholder[lclist[i]][beamzeroindex+(elementlist[j]-1)*6+1,1]*c;
                    rotation=resultsholder[lclist[i]][beamzeroindex+(elementlist[j]-1)*6+2,1];
                    isvalid=false;
                    if (beamloadsholder.ContainsKey(lclist[i]))
                    {
                        if (beamloadsholder[lclist[i]].ContainsKey(elementlist[j]))
                        {
                            isvalid=true;
                        }
                    }
                    if (isvalid)
                    {
                        cha=beamchapoints[lclist[i]][elementlist[j]];
                        bloads=beamloadsholder[lclist[i]][elementlist[j]];
                        effectmatrix=stiffnessmatcalcs.beamactioncalc(bloads,cha,axial,bending,rotation,resultsholder[lclist[i]][beamzeroindex+(elementlist[j]-1)*6+5,1]);

                    }
                    else
                    {
                        cha=new SortedSet<double> {0,1};
                        effectmatrix=Matrix<double>.Build.DenseOfArray(new double[,] {{-axial,bending,-rotation},{-axial,bending,resultsholder[lclist[i]][beamzeroindex+(elementlist[j]-1)*6+5,1]}});
                    }
                    for (int k = 0; k < effectmatrix.RowCount; k++)
                    {
                        if (cha.ElementAt(k) < 0.0001)
                        {
                            tempcha=0;
                        }
                        else
                        {
                            tempcha=cha.ElementAt(k);
                        }
                        tempstring=tempstring+string.Format("{0:G4}", tempcha)+"|"+string.Format("{0:G4}", tempcha*L)+"#";
                        for (int l = 0; l < 3; l++)
                        {
                            moveholder=effectmatrix[k,l];
                            if (Math.Abs(moveholder) < Math.Pow(10, -12))
                            {
                                moveholder=0;
                            }
                            tempstring=tempstring+string.Format("{0:G4}",moveholder)+",";
                        }
                        tempstring=tempstring.Substring(0,tempstring.Length-1)+"+";
                    }
                    outstring+=tempstring;
                    outstring=outstring.Substring(0,outstring.Length-1)+";";
                }
                outstring=outstring.Substring(0,outstring.Length-1)+":";
            }
            outstring=outstring.Substring(0,outstring.Length-1);
            
            return outstring;
        }
        public static Matrix<double> beamactioncalc(List<double[,]> beamloads, SortedSet<double> chapoints, double axial,double bending,double rotation,double rotation2)
        {
            var M=Matrix<double>.Build;
            Matrix<double> effectsholder=M.Dense(chapoints.Count,3);
            Matrix<double> holder=M.Dense(5,5);
            for (int i = 0; i < chapoints.Count; i++)
            {
                effectsholder[i,0]=-axial;
                effectsholder[i,1]=bending;
                effectsholder[i,2]=-rotation+(rotation2+rotation)*chapoints.ElementAt(i);
            }

            for (int i = 0; i < beamloads.Count; i++)
            {
                for (int j = 0; j < beamloads[i].GetLength(0); j++)
                {
                    switch (beamloads[i][j, 1])
                    {
                        case 0:
                            holder=elementactionsdisps.axialactionsdisps(beamloads[i],j,0,chapoints);
                            break;
                        case 1:
                            holder=elementactionsdisps.bendingactionsdisps(beamloads[i],j,0,chapoints);
                            break;
                        case 2:
                            holder=elementactionsdisps.bendingmomentactionsdisps(beamloads[i],j,0,chapoints);
                            break;
                    }
                    effectsholder=effectsholder.Add(holder);
                    
                }
            }


            return effectsholder;
        }
        public static Matrix<double> beamdispcalc(List<double[,]> beamloads, SortedSet<double> chapoints, double f0, double u0, double v0, double m0, double s0,double d0,double EA,double EI,double L)
        {
            var M=Matrix<double>.Build;
            Matrix<double> effectsholder=M.Dense(chapoints.Count,3);
            Matrix<double> holder=M.Dense(5,5);
            double x;
            double dirn=1;
            for (int i = 0; i < chapoints.Count; i++)
            {
                x=chapoints.ElementAt(i)*L;
                effectsholder[i,0]=dirn*((f0*x)/EA+u0);
                effectsholder[i,1]=dirn*((v0*Math.Pow(x,3)/6+m0*Math.Pow(x,2)/2)/EI+s0*x+d0);
                effectsholder[i,2]=dirn*((v0*Math.Pow(x,2)/2+m0*x)/EI+s0);
            }

            for (int i = 0; i < beamloads.Count; i++)
            {
                for (int j = 0; j < beamloads[i].GetLength(0); j++)
                {
                    switch (beamloads[i][j, 1])
                    {
                        case 0:
                            holder=elementactionsdisps.axialactionsdisps(beamloads[i],j,1,chapoints, EA);
                            break;
                        case 1:
                            holder=elementactionsdisps.bendingactionsdisps(beamloads[i],j,1,chapoints,EI);
                            break;
                        case 2:
                            holder=elementactionsdisps.bendingmomentactionsdisps(beamloads[i],j,1,chapoints,EI);
                            break;
                    }
                    effectsholder=effectsholder.Add(holder);
                    holder=M.Dense(chapoints.Count,3);
                    
                }
            }


            return effectsholder;
        }
    }
    class fixedendcreator
    {
        public static double[,] bendingloads(double[] bloadsrow,double[] beamgeomrow, double projectionscalar,double bendfraction)
        {
            double[,] outp= new double[1,13];
            outp[0,0]=bloadsrow[2];
            outp[0,1]=1;
            outp[0,2]=beamgeomrow[0];
            outp[0,3]=bloadsrow[5];
            outp[0,4]=bloadsrow[6]*projectionscalar*bendfraction;
            outp[0,5]=bloadsrow[7];
            outp[0,6]=bloadsrow[8]*projectionscalar*bendfraction;
            if (bloadsrow[2] == 0)
            {
                double P=bloadsrow[6]*projectionscalar*bendfraction;
                double a=beamgeomrow[0]*bloadsrow[5];
                double L=beamgeomrow[0];
                outp[0,8]=P * Math.Pow(L-a,2)*(L+2*a)/Math.Pow(L,3);
                outp[0,9]=P*a*Math.Pow(L-a,2)/Math.Pow(L,2);
                outp[0,11]=P * Math.Pow(a,2)*(L+2*(L-a))/Math.Pow(L,3);
                outp[0,12]=-P*(L-a)*Math.Pow(a,2)/Math.Pow(L,2);
            }
            else
            {
                double L=beamgeomrow[0];
                double wa=bloadsrow[6]*projectionscalar*bendfraction;
                double a=beamgeomrow[0]*bloadsrow[5];
                double wb=bloadsrow[8]*projectionscalar*bendfraction;
                double b=L-beamgeomrow[0]*bloadsrow[7];
                double Lw=L-a-b;
                double s1=10*((L*L+a*a)*(L+a)-(a*a+b*b)*(a-b)-L*b*(L+b)-a*a*a);
                double s2=Lw*(L*(2*L+a+b)-3*Math.Pow(a-b,2)-2*a*b);
                double s3=120*a*b*(a+Lw)+10*Lw*(6*a*a+4*L*Lw-3*Lw*Lw);
                double s4=10*L*Lw-10*Lw*a*(L-3*b)-9*Lw*Lw*Lw;
                double wm=wa/2+wb/2;
                double wd=wb-wa;
                outp[0,8]=Lw*wm-Lw*(s1*wm+s2*wd)/(20*Math.Pow(L,3));
                outp[0,9]=-(-Lw*(s3*wm+s4*wd)/(120*L*L)+Lw*(s1*wm+s2*wd)/(20*Math.Pow(L,3))*L-a*Lw*wm-Lw*Lw*(2*wb+wa)/6);
                outp[0,11]=Lw*(s1*wm+s2*wd)/(20*Math.Pow(L,3));
                outp[0,12]=-Lw*(s3*wm+s4*wd)/(120*L*L);

            }
            return outp;
        }
        public static double[,] bendingmomentloads(double[] bloadsrow,double[] beamgeomrow)
        {
            double[,] outp= new double[1,13];
            outp[0,0]=bloadsrow[2];
            outp[0,1]=2;
            outp[0,2]=beamgeomrow[0];
            outp[0,3]=bloadsrow[5];
            outp[0,4]=bloadsrow[6];
            outp[0,5]=bloadsrow[7];
            outp[0,6]=bloadsrow[8];
            if (bloadsrow[2] == 0)
            {
                double M=bloadsrow[6];
                double a=beamgeomrow[0]*bloadsrow[5];
                double L=beamgeomrow[0];
                outp[0,8]=-6*M*a*(L-a)/(L*L*L);
                outp[0,9]=M*(L-a)*(L-3*a)/(L*L);
                outp[0,11]=6*M*a*(L-a)/(L*L*L);
                outp[0,12]=M*a*(L-3*(L-a))/(L*L);
            }
            else
            {
                double ma=bloadsrow[6];
                double a=beamgeomrow[0]*bloadsrow[5];
                double mb=bloadsrow[8];
                double b=beamgeomrow[0]*bloadsrow[7];
                double L=beamgeomrow[0];
                double k=(mb-ma)/(b-a);
                double v0=(-4*L*a*a*ma - 2*L*a*a*mb + 2*L*a*b*ma - 2*L*a*b*mb + 2*L*b*b*ma + 4*L*b*b*mb + 3*a*a*a*ma + a*a*a*mb - a*a*b*ma + a*a*b*mb - a*b*b*ma + a*b*b*mb - b*b*b*ma - 3*b*b*b*mb)/(2*L*L*L);
                double c1=(-6*L*L*a*ma - 6*L*L*a*mb + 6*L*L*b*ma + 6*L*L*b*mb + 16*L*a*a*ma + 8*L*a*a*mb - 8*L*a*b*ma + 8*L*a*b*mb - 8*L*b*b*ma - 16*L*b*b*mb - 9*a*a*a*ma - 3*a*a*a*mb + 3*a*a*b*ma - 3*a*a*b*mb + 3*a*b*b*ma - 3*a*b*b*mb + 3*b*b*b*ma + 9*b*b*b*mb)/(12*L*L);
                outp[0,8]=-v0;
                outp[0,9]=c1;
                outp[0,11]=v0;
                outp[0,12]=-(v0 * a + c1+ (v0 - ma) * (b - a)- k/2 * Math.Pow(b - a,2)+ v0 * (L - b));
            }
            return outp;
        }
        public static double[,] axialloads(double[] bloadsrow,double[] beamgeomrow, double projectionscalar,double axfraction)
        {
           double[,] outp= new double[1,13];
            outp[0,0]=bloadsrow[2];
            outp[0,1]=0;
            outp[0,2]=beamgeomrow[0];
            outp[0,3]=bloadsrow[5];
            outp[0,4]=bloadsrow[6]*projectionscalar*axfraction;
            outp[0,5]=bloadsrow[7];
            outp[0,6]=bloadsrow[8]*projectionscalar*axfraction;
            if (bloadsrow[2] == 0)
            {
                double P=bloadsrow[6]*projectionscalar*axfraction;
                double a=beamgeomrow[0]*bloadsrow[5];
                double L=beamgeomrow[0];
                outp[0,7]=(1-a/L)*P;
                outp[0,10]=(a/L)*P;
            }
            else
            {
                double wa=bloadsrow[6]*projectionscalar*axfraction;
                double a=beamgeomrow[0]*bloadsrow[5];
                double wb=bloadsrow[8]*projectionscalar*axfraction;
                double b=beamgeomrow[0]*bloadsrow[7];
                double L=beamgeomrow[0];
                double k=(wb-wa)/(b-a);
                double Lt=wa*(b-a)+k*(b*b/2+a*a/2-a*b);
                outp[0,7]=Lt-(b*b*wa/2+k*b*b*b/3-k*b*b*a/2-a*a*wa/2+k*a*a*a/6)/L;
                outp[0,10]=(b*b*wa/2+k*b*b*b/3-k*b*b*a/2-a*a*wa/2+k*a*a*a/6)/L;
            }
            return outp;
        }
    }
    class elementactionsdisps
    {
        public static Matrix<double> axialactionsdisps(double[,] beamload, int rw, int actionsordisps,SortedSet<double> chapoints,double EA=1)
        {
            var M=Matrix<double>.Build;
            Matrix<double> effectsholder=M.Dense(chapoints.Count,3);
            for (int i = 0; i < chapoints.Count; i++)
            {
                double x=chapoints.ElementAt(i)*beamload[rw,2];
                if (beamload[rw,0] == 0)
                {
                    double P=beamload[rw,4];
                    double a=beamload[rw,2]*beamload[rw,3];
                    double L=beamload[rw,2];
                    if (actionsordisps == 0)
                    {
                        if (x < a)
                        {
                            effectsholder[i,0]=(1-a/L)*P;
                        }
                        else if (x>a)
                        {
                            effectsholder[i,0]=-(a/L)*P;
                        }
                    }
                    else
                    {
                        if (x <= a)
                        {
                            effectsholder[i,0]=(1-a/L)*P*x/EA;
                        }
                        else if (x>a)
                        {
                            effectsholder[i,0]=(1-a/L)*P*a/EA-(a/L)*P*(x-a)/EA;
                        }
                    }
                    

                }
                else
                {
                    double wa=beamload[rw,4];
                    double a=beamload[rw,2]*beamload[rw,3];
                    double wb=beamload[rw,6];
                    double b=beamload[rw,2]*beamload[rw,5];
                    double L=beamload[rw,2];
                    double k=(wb-wa)/(b-a);
                    double Lt=wa*(b-a)+k*(b*b/2+a*a/2-a*b);
                    if (actionsordisps == 0)
                    {
                        if (x < a)
                        {
                            effectsholder[i,0]=Lt-(b*b*wa/2+k*b*b*b/3-k*b*b*a/2-a*a*wa/2+k*a*a*a/6)/L;
                        }
                        else if (x<b)
                        {
                            effectsholder[i,0]=(Lt-(b*b*wa/2+k*b*b*b/3-k*b*b*a/2-a*a*wa/2+k*a*a*a/6)/L)-(wa*x+k*x*x/2-k*a*x-a*wa+k*a*a/2);
                        }
                        else
                        {
                            effectsholder[i,0]=-(b*b*wa/2+k*b*b*b/3-k*b*b*a/2-a*a*wa/2+k*a*a*a/6)/L;
                        }
                    }
                    else
                    {
                        if (x < a)
                        {
                            effectsholder[i,0]=(Lt-(b*b*wa/2+k*b*b*b/3-k*b*b*a/2-a*a*wa/2+k*a*a*a/6)/L)*x/EA;
                        }
                        else if (x<b)
                        {
                            double comp1=(Lt-(b*b*wa/2+k*b*b*b/3-k*b*b*a/2-a*a*wa/2+k*a*a*a/6)/L)*a/EA;
                            double comp2=(Lt-(b*b*wa/2+k*b*b*b/3-k*b*b*a/2-a*a*wa/2+k*a*a*a/6)/L)*(x-a)/EA;
                            double comp3=(-(wa*x*x/2+k*x*x*x/6-k*a*x*x/2-a*wa*x+k*a*a*x/2)+(wa*a*a/2+k*a*a*a/6-k*a*a*a/2-a*wa*a+k*a*a*a/2))/EA;
                            effectsholder[i,0]=comp1+comp2+comp3;
                        }
                        else
                        {
                            effectsholder[i,0]=-(b*b*wa/2+k*b*b*b/3-k*b*b*a/2-a*a*wa/2+k*a*a*a/6)/L*(x-L)/EA;
                        }
                    }
                }
            }
            
            return effectsholder;
        }
        public static Matrix<double> bendingactionsdisps(double[,] beamload, int rw, int actionsordisps,SortedSet<double> chapoints,double EI=1)
        {
           var M=Matrix<double>.Build;
            Matrix<double> effectsholder=M.Dense(chapoints.Count,3);
            for (int i = 0; i < chapoints.Count; i++)
            {
                double x=chapoints.ElementAt(i)*beamload[rw,2];
                if (beamload[rw,0] == 0)
                {
                    double P=beamload[rw,4];
                    double a=beamload[rw,2]*beamload[rw,3];
                    double L=beamload[rw,2];
                    double b=L-a;
                    if (actionsordisps == 0)
                    {
                        if (x < a)
                        {
                            effectsholder[i,1]=-P * Math.Pow(L-a,2)*(L+2*a)/Math.Pow(L,3);
                            effectsholder[i,2]=P*b*b*(a*L-(L+2*a)*x)/Math.Pow(L,3);
                        }
                        else if (x>a)
                        {
                            effectsholder[i,1]=P * Math.Pow(a,2)*(L+2*(L-a))/Math.Pow(L,3);
                            effectsholder[i,2]=-P*a*a*(L*L+b*L-(L+2*b)*x)/Math.Pow(L,3);
                        }
                    }
                    else
                    {
                        double RA=P*b*b*(L+2*a)/Math.Pow(L,3);
                        double MA=-P*a*b*b/Math.Pow(L,2);
                        if (x <= a)
                        {
                            effectsholder[i,1]=-RA*Math.Pow(x,3)/6/EI-MA*Math.Pow(x,2)/2/EI;
                            effectsholder[i,2]=-RA*Math.Pow(x,2)/2/EI-MA*x/EI;
                        }
                        else if (x>a)
                        {
                            effectsholder[i,1]=-RA*Math.Pow(x,3)/6/EI-MA*Math.Pow(x,2)/2/EI+P*Math.Pow(x-a,3)/6/EI;
                            effectsholder[i,2]=-RA*Math.Pow(x,2)/2/EI-MA*x/EI+P*Math.Pow(x-a,2)/2/EI;
                        }
                    }
                    

                }
                else
                {
                    double L=beamload[rw,2];
                    double wa=beamload[rw,4];
                    double a=beamload[rw,2]*beamload[rw,3];
                    double wb=beamload[rw,6];
                    double b=L-beamload[rw,2]*beamload[rw,5];
                    double Lw=L-a-b;
                    double s1=10*((L*L+a*a)*(L+a)-(a*a+b*b)*(a-b)-L*b*(L+b)-a*a*a);
                    double s2=Lw*(L*(2*L+a+b)-3*Math.Pow(a-b,2)-2*a*b);
                    double s3=120*a*b*(a+Lw)+10*Lw*(6*a*a+4*L*Lw-3*Lw*Lw);
                    double s4=10*L*Lw-10*Lw*a*(L-3*b)-9*Lw*Lw*Lw;
                    double wm=wa/2+wb/2;
                    double wd=wb-wa;
                    double Ra=-(Lw*wm-Lw*(s1*wm+s2*wd)/(20*Math.Pow(L,3)));
                    double Rb=Lw*(s1*wm+s2*wd)/(20*Math.Pow(L,3));
                    double Ma=-(-Lw*(s3*wm+s4*wd)/(120*L*L)+Lw*(s1*wm+s2*wd)/(20*Math.Pow(L,3))*L-a*Lw*wm-Lw*Lw*(2*wb+wa)/6);
                    double Mb=-Lw*(s3*wm+s4*wd)/(120*L*L);
                    double wx=wa+(x-a)*(wb-wa)/Lw;
                    if (actionsordisps == 0)
                    {
                        if (x < a)
                        {
                            effectsholder[i,1]=Ra;
                            effectsholder[i,2]=Ma+Ra*x;
                        }
                        else if (x<(L-b))
                        {
                            effectsholder[i,1]=Ra+(wa+wx)*(x-a)/2;
                            effectsholder[i,2]=Ma+Ra*x+(2*wa+wx)*Math.Pow(x-a,2)/6;
                        }
                        else
                        {
                            effectsholder[i,1]=Rb;
                            effectsholder[i,2]=-Mb-(L-x)*Rb;
                        }
                    }
                    else
                    {
                        if (x < a)
                        {
                            effectsholder[i,1]=Ra*Math.Pow(x,3)/6/EI+Ma*Math.Pow(x,2)/2/EI;
                            effectsholder[i,2]=Ra*Math.Pow(x,2)/2/EI+Ma*x/EI;
                        }
                        else if (x<(L-b))
                        {
                            effectsholder[i,1]=Ra*Math.Pow(x,3)/6/EI+Ma*Math.Pow(x,2)/2/EI+(4*wa+wx)*Math.Pow(x-a,4)/120/EI;
                            effectsholder[i,2]=Ra*Math.Pow(x,2)/2/EI+Ma*x/EI+(3*wa+wx)*Math.Pow(x-a,3)/24/EI;
                        }
                        else
                        {
                            effectsholder[i,1]=-Rb*Math.Pow(L-x,3)/6/EI-Mb*Math.Pow(L-x,2)/2/EI;
                            effectsholder[i,2]=+Rb*Math.Pow(L-x,2)/2/EI+Mb*(L-x)/EI;
                        }
                    }
                }
            }
            
            return effectsholder;
        }
        public static Matrix<double> bendingmomentactionsdisps(double[,] beamload, int rw, int actionsordisps,SortedSet<double> chapoints,double EI=1)
        {
            var M=Matrix<double>.Build;
            Matrix<double> effectsholder=M.Dense(chapoints.Count,3);
            for (int i = 0; i < chapoints.Count; i++)
            {
                double x=chapoints.ElementAt(i)*beamload[rw,2];
                if (beamload[rw,0] == 0)
                {
                    double P=beamload[rw,4];
                    double a=beamload[rw,2]*beamload[rw,3];
                    double L=beamload[rw,2];
                    double b=L-a;
                    double Ra=6*P*a*b/Math.Pow(L,3);
                    double Ma=P*b*(L-3*a)/Math.Pow(L,2);
                    if (actionsordisps== 0)
                    {
                        if (x < a)
                        {
                            effectsholder[i,1]=6*P*a*(L-a)/(L*L*L);
                            effectsholder[i,2]=P*b*(L*(L-3*a)+6*a*x)/Math.Pow(L,3);
                        }
                        else if (x>a)
                        {
                            effectsholder[i,1]=6*P*a*(L-a)/(L*L*L);
                            effectsholder[i,2]=P*b*(L*(L-3*a)+6*a*x)/Math.Pow(L,3)-P;
                        }
                    }
                    else
                    {
                        if (x <= a)
                        {
                            effectsholder[i,1]=Ra*Math.Pow(x,3)/6/EI+Ma*Math.Pow(x,2)/2/EI;
                            effectsholder[i,2]=Ra*Math.Pow(x,2)/2/EI+Ma*x/EI;
                        }
                        else if (x>a)
                        {
                            effectsholder[i,1]=Ra*Math.Pow(x,3)/6/EI+Ma*Math.Pow(x,2)/2/EI-P*Math.Pow(x-a,2)/2/EI;
                            effectsholder[i,2]=Ra*Math.Pow(x,2)/2/EI+Ma*x/EI-P*(x-a)/EI;
                        }
                    }

                }
                else
                {
                    double L=beamload[rw,2];
                    double ma=beamload[rw,4];
                    double a=beamload[rw,2]*beamload[rw,3];
                    double mb=beamload[rw,6];
                    double b=beamload[rw,2]*beamload[rw,5];
                    double k=(mb-ma)/(b-a);
                    double v0=(-4*L*a*a*ma - 2*L*a*a*mb + 2*L*a*b*ma - 2*L*a*b*mb + 2*L*b*b*ma + 4*L*b*b*mb + 3*a*a*a*ma + a*a*a*mb - a*a*b*ma + a*a*b*mb - a*b*b*ma + a*b*b*mb - b*b*b*ma - 3*b*b*b*mb)/(2*L*L*L);
                    double c1=(-6*L*L*a*ma - 6*L*L*a*mb + 6*L*L*b*ma + 6*L*L*b*mb + 16*L*a*a*ma + 8*L*a*a*mb - 8*L*a*b*ma + 8*L*a*b*mb - 8*L*b*b*ma - 16*L*b*b*mb - 9*a*a*a*ma - 3*a*a*a*mb + 3*a*a*b*ma - 3*a*a*b*mb + 3*a*b*b*ma - 3*a*b*b*mb + 3*b*b*b*ma + 9*b*b*b*mb)/(12*L*L);
                    double d2 =
                        a * (
                        -6*Math.Pow(L,3)*a*ma
                        -6*Math.Pow(L,3)*a*mb
                        +6*Math.Pow(L,3)*b*ma
                        +6*Math.Pow(L,3)*b*mb
                        +16*Math.Pow(L,2)*Math.Pow(a,2)*ma
                        +8*Math.Pow(L,2)*Math.Pow(a,2)*mb
                        -8*Math.Pow(L,2)*a*b*ma
                        +8*Math.Pow(L,2)*a*b*mb
                        -8*Math.Pow(L,2)*Math.Pow(b,2)*ma
                        -16*Math.Pow(L,2)*Math.Pow(b,2)*mb
                        -21*L*Math.Pow(a,3)*ma
                        -9*L*Math.Pow(a,3)*mb
                        +9*L*Math.Pow(a,2)*b*ma
                        -9*L*Math.Pow(a,2)*b*mb
                        +9*L*a*Math.Pow(b,2)*ma
                        +9*L*a*Math.Pow(b,2)*mb
                        +3*L*Math.Pow(b,3)*ma
                        +9*L*Math.Pow(b,3)*mb
                        +9*Math.Pow(a,4)*ma
                        +3*Math.Pow(a,4)*mb
                        -3*Math.Pow(a,3)*b*ma
                        +3*Math.Pow(a,3)*b*mb
                        -3*Math.Pow(a,2)*Math.Pow(b,2)*ma
                        +3*Math.Pow(a,2)*Math.Pow(b,2)*mb
                        -3*a*Math.Pow(b,3)*ma
                        -9*a*Math.Pow(b,3)*mb
                    ) / (12*Math.Pow(L,3));

                    double d3 =
                    Math.Pow(a,2) * (
                        -6*Math.Pow(L,3)*a*ma
                        -6*Math.Pow(L,3)*a*mb
                        +6*Math.Pow(L,3)*b*ma
                        +6*Math.Pow(L,3)*b*mb
                        +16*Math.Pow(L,2)*Math.Pow(a,2)*ma
                        +8*Math.Pow(L,2)*Math.Pow(a,2)*mb
                        -8*Math.Pow(L,2)*a*b*ma
                        +8*Math.Pow(L,2)*a*b*mb
                        -8*Math.Pow(L,2)*Math.Pow(b,2)*ma
                        -16*Math.Pow(L,2)*Math.Pow(b,2)*mb
                        -17*L*Math.Pow(a,3)*ma
                        -7*L*Math.Pow(a,3)*mb
                        +7*L*Math.Pow(a,2)*b*ma
                        -7*L*Math.Pow(a,2)*b*mb
                        +7*L*a*Math.Pow(b,2)*ma
                        +5*L*a*Math.Pow(b,2)*mb
                        +3*L*Math.Pow(b,3)*ma
                        +9*L*Math.Pow(b,3)*mb
                        +6*Math.Pow(a,4)*ma
                        +2*Math.Pow(a,4)*mb
                        -2*Math.Pow(a,3)*b*ma
                        +2*Math.Pow(a,3)*b*mb
                        -2*Math.Pow(a,2)*Math.Pow(b,2)*ma
                        +2*Math.Pow(a,2)*Math.Pow(b,2)*mb
                        -2*a*Math.Pow(b,3)*ma
                        -6*a*Math.Pow(b,3)*mb
                    ) / (24*Math.Pow(L,3));

                    double e2 =
                    (
                        -4*Math.Pow(L,3)*Math.Pow(a,2)*ma
                        -2*Math.Pow(L,3)*Math.Pow(a,2)*mb
                        +2*Math.Pow(L,3)*a*b*ma
                        -2*Math.Pow(L,3)*a*b*mb
                        +2*Math.Pow(L,3)*Math.Pow(b,2)*ma
                        +4*Math.Pow(L,3)*Math.Pow(b,2)*mb
                        +16*Math.Pow(L,2)*Math.Pow(a,2)*b*ma
                        +8*Math.Pow(L,2)*Math.Pow(a,2)*b*mb
                        -8*Math.Pow(L,2)*a*Math.Pow(b,2)*ma
                        +8*Math.Pow(L,2)*a*Math.Pow(b,2)*mb
                        -8*Math.Pow(L,2)*Math.Pow(b,3)*ma
                        -16*Math.Pow(L,2)*Math.Pow(b,3)*mb
                        -9*L*Math.Pow(a,3)*b*ma
                        -3*L*Math.Pow(a,3)*b*mb
                        -9*L*Math.Pow(a,2)*Math.Pow(b,2)*ma
                        -9*L*Math.Pow(a,2)*Math.Pow(b,2)*mb
                        +9*L*a*Math.Pow(b,3)*ma
                        -9*L*a*Math.Pow(b,3)*mb
                        +9*L*Math.Pow(b,4)*ma
                        +21*L*Math.Pow(b,4)*mb
                        +9*Math.Pow(a,3)*Math.Pow(b,2)*ma
                        +3*Math.Pow(a,3)*Math.Pow(b,2)*mb
                        -3*Math.Pow(a,2)*Math.Pow(b,3)*ma
                        +3*Math.Pow(a,2)*Math.Pow(b,3)*mb
                        -3*a*Math.Pow(b,4)*ma
                        +3*a*Math.Pow(b,4)*mb
                        -3*Math.Pow(b,5)*ma
                        -9*Math.Pow(b,5)*mb
                    ) / (12*Math.Pow(L,3));

                    double e3 =
                    (
                        3*Math.Pow(L,3)*Math.Pow(a,3)*ma
                        +Math.Pow(L,3)*Math.Pow(a,3)*mb
                        -9*Math.Pow(L,3)*Math.Pow(a,2)*b*ma
                        -3*Math.Pow(L,3)*Math.Pow(a,2)*b*mb
                        +3*Math.Pow(L,3)*a*Math.Pow(b,2)*ma
                        -3*Math.Pow(L,3)*a*Math.Pow(b,2)*mb
                        +3*Math.Pow(L,3)*Math.Pow(b,3)*ma
                        +5*Math.Pow(L,3)*Math.Pow(b,3)*mb
                        +16*Math.Pow(L,2)*Math.Pow(a,2)*Math.Pow(b,2)*ma
                        +8*Math.Pow(L,2)*Math.Pow(a,2)*Math.Pow(b,2)*mb
                        -8*Math.Pow(L,2)*a*Math.Pow(b,3)*ma
                        +8*Math.Pow(L,2)*a*Math.Pow(b,3)*mb
                        -8*Math.Pow(L,2)*Math.Pow(b,4)*ma
                        -16*Math.Pow(L,2)*Math.Pow(b,4)*mb
                        -9*L*Math.Pow(a,3)*Math.Pow(b,2)*ma
                        -3*L*Math.Pow(a,3)*Math.Pow(b,2)*mb
                        -5*L*Math.Pow(a,2)*Math.Pow(b,3)*ma
                        -7*L*Math.Pow(a,2)*Math.Pow(b,3)*mb
                        +7*L*a*Math.Pow(b,4)*ma
                        -7*L*a*Math.Pow(b,4)*mb
                        +7*L*Math.Pow(b,5)*ma
                        +17*L*Math.Pow(b,5)*mb
                        +6*Math.Pow(a,3)*Math.Pow(b,3)*ma
                        +2*Math.Pow(a,3)*Math.Pow(b,3)*mb
                        -2*Math.Pow(a,2)*Math.Pow(b,4)*ma
                        +2*Math.Pow(a,2)*Math.Pow(b,4)*mb
                        -2*a*Math.Pow(b,5)*ma
                        +2*a*Math.Pow(b,5)*mb
                        -2*Math.Pow(b,6)*ma
                        -6*Math.Pow(b,6)*mb
                    ) / (24*Math.Pow(L,3));

                    double m2b=(v0*a+c1+(v0-ma)*(b-a)-k/2*Math.Pow(b-a,2));
                    if (actionsordisps== 0)
                    {
                        if (x < a)
                        {
                            effectsholder[i,1]=v0;
                            effectsholder[i,2]=v0*x+c1;
                        }
                        else if (x<b)
                        {
                            effectsholder[i,1]=v0;
                            effectsholder[i,2]=v0 * a + c1+ (v0 - ma) * (x - a)- k/2 * Math.Pow(x - a,2);
                        }
                        else
                        {
                            effectsholder[i,1]=v0;
                            effectsholder[i,2]=v0 * a + c1+ (v0 - ma) * (b - a)- k/2 * Math.Pow(b - a,2) +v0*(x-b);
                        }
                    }
                    else
                    {
                        if (x < a)
                        {
                            effectsholder[i,1]=(v0*Math.Pow(x,3)/6+c1*Math.Pow(x,2)/2)/EI;
                            effectsholder[i,2]=(v0*Math.Pow(x,2)/2+c1*x)/EI;
                        }
                        else if (x<b)
                        {
                            effectsholder[i,1]= ((v0 * a + c1) * Math.Pow(x - a,2) / 2+ (v0 - ma) * Math.Pow(x - a,3) / 6- k/24 * Math.Pow(x - a,4)+ d2 * (x - a)+ d3)/(EI);
                            effectsholder[i,2]= ((v0 * a + c1) * (x - a)+ (v0 - ma) * Math.Pow(x - a,2) / 2- k/6 * Math.Pow(x - a,3)+ d2)/(EI);
                        }
                        else
                        {
                            effectsholder[i,1]=(m2b*Math.Pow(x-b,2)/2+v0*Math.Pow(x-b,3)/6+e2*(x-b)+e3)/EI;
                            effectsholder[i,2]=(m2b*(x-b)+v0*Math.Pow(x-b,2)/2+e2)/EI;
                        }
                    }
                }
            }
            
            return effectsholder;
        }
    }
   
    class interfacefunctions
    {
        public static object[,] filterarrayempties(object[,] inarray)
        {
            List<int> valids= new List<int>();
            for (int i = 0; i < inarray.GetLength(0); i++)
            {
                if (!(inarray[i,0] is ExcelEmpty))
                {
                    valids.Add(i);
                }
            }
            object[,] outarray= new object[valids.Count,inarray.GetLength(1)];
                int rowit=0;
                foreach (int valid in valids)
                {
                    for (int j = 0; j < inarray.GetLength(1); j++)
                    {
                        if (!(inarray[valid,j] is ExcelEmpty)&&!(inarray[valid,j] is ExcelMissing))
                        {
                            outarray[rowit,j]=inarray[valid,j];
                        }
                    }
                    rowit++;
                }
            return outarray;
        }
        public static double[,] nodefilter(object[,] inarray)
        {
            double[,] outarray=new double[inarray.GetLength(0),inarray.GetLength(1)];
            string[] temp;
            int iter;
            for (int i = 0; i < inarray.GetLength(0); i++)
            {
                outarray[i,0]=(double)inarray[i,0];
                outarray[i,1]=(double)inarray[i,1];
                for (int j = 0; j < 3; j++)
                {
                    if (inarray[i,2+j] is bool)
                    {
                        if ((bool)inarray[i, 2 + j])
                        {
                            outarray[i,2+j]=-1;
                        }
                        else
                        {
                             outarray[i,2+j]=0;
                        }
                    }
                    else if ((inarray[i,2+j].ToString() ?? string.Empty).Contains(";"))
                    {
                        temp=(inarray[i,2+j].ToString() ?? string.Empty).Split(";");
                        iter=0;
                        while (double.Parse(temp[iter].Split(",")[0]) * double.Parse(temp[iter + 1].Split(",")[0]) > 0)
                        {
                            iter++;
                        }
                        if (double.Parse(temp[iter].Split(",")[0]) * double.Parse(temp[iter + 1].Split(",")[0]) == 0)
                        {
                            outarray[i,2+j]=Math.Max((double.Parse(temp[iter + 1].Split(",")[1])-double.Parse(temp[iter].Split(",")[1]))/(double.Parse(temp[iter + 1].Split(",")[0])-double.Parse(temp[iter].Split(",")[0])),(double.Parse(temp[iter + 2].Split(",")[1])-double.Parse(temp[iter+1].Split(",")[1]))/(double.Parse(temp[iter + 2].Split(",")[0])-double.Parse(temp[iter+1].Split(",")[0])));
                        }
                        else
                        {
                            outarray[i,2+j]=(double.Parse(temp[iter + 1].Split(",")[1])-double.Parse(temp[iter].Split(",")[1]))/(double.Parse(temp[iter + 1].Split(",")[0])-double.Parse(temp[iter].Split(",")[0]));
                        }
                        
                    }

                    else
                    {
                        outarray[i,2+j]=(double)inarray[i,2+j];
                    }
                }
            }

            return outarray;
        }
        public static double[,] elementfilter(object[,] inarray)
        {
            double[,] outarray=new double[inarray.GetLength(0),inarray.GetLength(1)];
            for (int i = 0; i < inarray.GetLength(0); i++)
            {
                outarray[i,0]=(double)inarray[i,0];
                outarray[i,1]=(double)inarray[i,1];
                outarray[i,2]=(double)inarray[i,2];
                outarray[i,3]=(double)inarray[i,3];
                for (int j = 0; j < 6; j++)
                {
                    if ((bool)inarray[i, 4 + j])
                    {
                        outarray[i,4+j]=1;
                    }
                    else
                    {
                        outarray[i,4+j]=0;
                    }
                }
            }

            return outarray;
        }
        public static double[,] nloadsfilter(object[,] inarray, double[,] tnodes)
        {
            
            List<int> memberlist;
            List<int> inlist= new List<int>();
            List<double[]> outlist=new List<double[]>();
            double[] temparray;
            string memberstring;
            for (int i = 1; i <= tnodes.GetLength(0);i++)
            {
                inlist.Add(i);
            }
            for (int i = 0; i < inarray.GetLength(0); i++)
            {
                memberstring=inarray[i,1].ToString()?? string.Empty;
                if (memberstring.ToLower() == "all")
                {
                    memberstring="-1";
                }
                memberlist = stiffnessmatcalcs.lcmemberstringread(memberstring, tnodes.GetLength(0),inlist);
                foreach(int member in memberlist)
                {
                    temparray=new double[4];
                    temparray[0]=(double)inarray[i,0];
                    temparray[1]=(double)member;
                    switch (inarray[i, 2])
                    {
                        case "x":
                            temparray[2]=0;
                            break;
                        case "y":
                            temparray[2]=1;
                            break;
                        case "zz":
                            temparray[2]=2;
                            break;
                    }
                    if ((inarray[i, 3].ToString()?? string.Empty).Substring(0,1) == "#")
                    {
                        temparray[3]=parseclass.parseloadequations(temparray[0],temparray[1],tnodes[(int)temparray[1]-1,0],tnodes[(int)temparray[1]-1,1],0,(inarray[i, 3].ToString()?? string.Empty).Substring(1));
                    }
                    else
                    {
                        temparray[3]=(double)inarray[i,3];
                    }
                    outlist.Add(temparray);
                }
                
            }
            double[,] outarray=new double[outlist.Count,inarray.GetLength(1)];
            for (int i = 0; i < outlist.Count; i++)
            {
                if (outlist[i][1]>tnodes.GetLength(0) || outlist[i][1] < 1)
                {
                    throw new ArgumentNullException("Provided node index is out of range");
                }
                for (int j = 0; j < inarray.GetLength(1); j++)
                {
                    outarray[i,j]=outlist[i][j];
                }
            }

            return outarray;
        }
        public static double[,] bloadsfilter(object[,] inarray,double[,] telements,double[,] tnodes)
        {
            List<int> memberlist;
            List<int> inlist= new List<int>();
            List<double[]> outlist=new List<double[]>();
            double[] temparray;
            string memberstring;
            for (int i = 1; i <= telements.GetLength(0);i++)
            {
                inlist.Add(i);
            }
            bool patchmarker;
            double LHSpercent;
            double LHSload;
            double RHSpercent;
            double RHSload;
            double x1;
            double y1;
            double x2;
            double y2;
            for (int i = 0; i < inarray.GetLength(0); i++)
            {
                memberstring=inarray[i,1].ToString()?? string.Empty;
                if (memberstring.ToLower() == "all")
                {
                    memberstring="-1";
                }
                memberlist = stiffnessmatcalcs.lcmemberstringread(memberstring, telements.GetLength(0),inlist);
                foreach(int member in memberlist)
                {
                    temparray=new double[9];
                    temparray[0]=(double)inarray[i,0];
                    temparray[1]=member;
                    if ((String)inarray[i,2]=="Point Load")
                    {
                        patchmarker=false;
                        temparray[2]=0;
                    }
                    else
                    {
                        patchmarker=true;
                        temparray[2]=1;
                    }
                    switch ((String)inarray[i,3])
                    {
                        case "Local":
                            temparray[3]=0;
                            break;
                        case "Global":
                            temparray[3]=1;
                            break;
                        case "Global Projected":
                            temparray[3]=2;
                            break;
                    }
                    switch (inarray[i, 4])
                    {
                        case "x":
                            temparray[4]=0;
                            break;
                        case "y":
                            temparray[4]=1;
                            break;
                        case "zz":
                            temparray[4]=2;
                            break;
                    }
                    if (patchmarker)
                    {
                        LHSpercent=(double)inarray[i,5];
                        
                        RHSpercent=(double)inarray[i,7];
                        
                        if ((inarray[i, 6].ToString()?? string.Empty).Substring(0,1) == "#")
                        {
                            x1=tnodes[(int)telements[(int)temparray[1]-1,0]-1,0];
                            y1=tnodes[(int)telements[(int)temparray[1]-1,0]-1,1];
                            x2=tnodes[(int)telements[(int)temparray[1]-1,1]-1,0];
                            y2=tnodes[(int)telements[(int)temparray[1]-1,1]-1,1];
                            if (LHSpercent < RHSpercent)
                            {
                                temparray[5]=LHSpercent;
                                temparray[6]=parseclass.parseloadequations(temparray[0],temparray[1],x1+LHSpercent*(x2-x1),y1+LHSpercent*(y2-y1),0,(inarray[i, 6].ToString()?? string.Empty).Substring(1));
                                temparray[7]=RHSpercent;
                                temparray[8]=parseclass.parseloadequations(temparray[0],temparray[1],x1+RHSpercent*(x2-x1),y1+RHSpercent*(y2-y1),0,(inarray[i, 6].ToString()?? string.Empty).Substring(1));
                            }
                            else
                            {
                                temparray[5]=RHSpercent;
                                temparray[6]=parseclass.parseloadequations(temparray[0],temparray[1],x1+RHSpercent*(x2-x1),y1+RHSpercent*(y2-y1),0,(inarray[i, 6].ToString()?? string.Empty).Substring(1));
                                temparray[7]=LHSpercent;
                                temparray[8]=parseclass.parseloadequations(temparray[0],temparray[1],x1+LHSpercent*(x2-x1),y1+LHSpercent*(y2-y1),0,(inarray[i, 6].ToString()?? string.Empty).Substring(1));
                            }
                        }
                        else
                        {
                            LHSload=(double)inarray[i,6];
                            RHSload=(double)inarray[i,8];
                            if (LHSpercent < RHSpercent)
                            {
                                temparray[5]=LHSpercent;
                                temparray[6]=LHSload;
                                temparray[7]=RHSpercent;
                                temparray[8]=RHSload;
                            }
                            else
                            {
                                temparray[5]=RHSpercent;
                                temparray[6]=RHSload;
                                temparray[7]=LHSpercent;
                                temparray[8]=LHSload;
                            }
                        }
                        
                    }
                    else
                    {
                        temparray[5]=(double)inarray[i,5];
                        if ((inarray[i, 6].ToString()?? string.Empty).Substring(0,1) == "#")
                        {
                            x1=tnodes[(int)telements[(int)temparray[1]-1,0]-1,0];
                            y1=tnodes[(int)telements[(int)temparray[1]-1,0]-1,1];
                            x2=tnodes[(int)telements[(int)temparray[1]-1,1]-1,0];
                            y2=tnodes[(int)telements[(int)temparray[1]-1,1]-1,1];
                            temparray[6]=parseclass.parseloadequations(temparray[0],temparray[1],x1+(double)inarray[i,5]*(x2-x1),y1+(double)inarray[i,5]*(y2-y1),0,(inarray[i, 6].ToString()?? string.Empty).Substring(1));
                        }
                        else
                        {
                            temparray[6]=(double)inarray[i,6];
                        }

                        temparray[7]=0;
                        temparray[8]=0;
                    }
                    outlist.Add(temparray);
                }
                
            }
            double[,] outarray=new double[outlist.Count,inarray.GetLength(1)];
            for (int i = 0; i < outlist.Count; i++)
            {
                if (outlist[i][1]>telements.GetLength(0) || outlist[i][1] < 1)
                {
                    throw new ArgumentNullException("Provided node index is out of range");
                }
                for (int j = 0; j < inarray.GetLength(1); j++)
                {
                    outarray[i,j]=outlist[i][j];
                }
            }

            return outarray;
        }
        public static object[,] extractsfilter(object[,] inarray)
        {
           object[,] outarray=new object[inarray.GetLength(0),inarray.GetLength(1)];
            for (int i = 0; i < inarray.GetLength(0); i++)
            {
                if ((inarray[i, 0].ToString()?? string.Empty).ToLower() == "all")
                {
                    outarray[i,0]=-1;
                }
                else
                {
                    outarray[i,0]=inarray[i, 0];
                }
                if ((String)inarray[i, 1] == "Actions")
                {
                    outarray[i, 1]=0;
                }
                else
                {
                    outarray[i, 1]=1;
                }
                if ((String)inarray[i, 2] == "Nodes")
                {
                    outarray[i, 2]=0;
                }
                else
                {
                    outarray[i, 2]=1;
                }
                
                if ((inarray[i, 3].ToString()?? string.Empty).ToLower() == "all")
                {
                    outarray[i,3]=-1;
                }
                else
                {
                    outarray[i,3]=inarray[i, 3];
                }
                
            }
            return outarray;
        }
        public static Dictionary<(int,int),List<double[]>> tnsprings(object[,] inarray)
        {
            List<double[]> templist;
            double[] rw;
            string[] tempstring;
            Dictionary<(int,int),List<double[]>> outdict=new Dictionary<(int,int),List<double[]>>();
            for (int i = 0; i < inarray.GetLength(0); i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    if ((inarray[i,2+j].ToString() ?? string.Empty).Contains(";"))
                    {
                        templist= new List<double[]>();
                        tempstring=(inarray[i,2+j].ToString() ?? string.Empty).Split(";");
                        for (int k = 0; k < tempstring.GetLength(0); k++)
                        {
                            rw= new double[] {double.Parse(tempstring[k].Split(",")[0]),double.Parse(tempstring[k].Split(",")[1])};
                            templist.Add(rw);
                        }
                        outdict.Add((i,j),templist);
                    }
                }
            }

            return outdict;
        }

    }
    class nonlinearfunctions
    {
        public static List<Vector<double>> nonlinspring(Vector<double> uhat,Vector<double> fhat, Matrix<double> kmat,Dictionary<(int,int),List<double[]>> tnsprings,List<int> springmap, double[,] tnodes,double[] filtermap)
        {
            var V=Vector<double>.Build;
            Vector<double> deltau=V.Dense(uhat.Count);
            Matrix<double> ktemp;
            Matrix<double> kreduced;
            Vector<double> ghatreduced;
            Vector<double> deltauunfilter;
            int itercount=0;
            List<Vector<double>> ghatcalc=nonlinearfunctions.residualsforcereversestiffness(uhat,fhat,tnsprings,springmap);
            while (ghatcalc[0].L1Norm()>0.01 || deltau.L1Norm() > 0.000001)
            {
                fhat=fhat-ghatcalc[0]+ghatcalc[1];
                ktemp=nonlinearfunctions.kcreate(kmat,tnsprings,springmap,tnodes,ghatcalc[2]);
                kreduced=stiffnessmatcalcs.kmatreduce(ktemp,filtermap);
                ghatreduced=stiffnessmatcalcs.filterby(ghatcalc[0],filtermap);
                deltau=(kreduced.Inverse()).Multiply(ghatreduced);
                deltauunfilter=stiffnessmatcalcs.unfilterby(deltau,filtermap,uhat.Count);
                uhat=uhat+deltauunfilter;
                fhat=fhat+ktemp.Multiply(deltauunfilter);
                ghatcalc=nonlinearfunctions.residualsforcereversestiffness(uhat,fhat,tnsprings,springmap);
                if (ghatcalc[0].L1Norm() == 0)
                {
                    break;
                }
                itercount=itercount+1;
                if (itercount > 50)
                {
                    throw new InvalidOperationException("Convergence not achieved in 50 iterations");
                }
            }
            List<Vector<double>> outlist= new List<Vector<double>>();
            outlist.Add(uhat);
            outlist.Add(fhat);
            return outlist;
        }
        public static Matrix<double> kcreate(Matrix<double> kmatreduce,Dictionary<(int,int),List<double[]>> tnsprings,List<int> springmap, double[,] tnodes,Vector<double> stiff)
        {
            Matrix<double> kmatnew=kmatreduce.Clone();
            int dofindex;
            double initstif;
            double newstif;
            double deltastif;
            foreach((int,int) key in tnsprings.Keys)
            {
                dofindex=springmap[key.Item1*3+key.Item2];
                initstif=tnodes[key.Item1,key.Item2+2];
                newstif=stiff[dofindex];
                deltastif=newstif-initstif;
                kmatnew[key.Item1*3+key.Item2,key.Item1*3+key.Item2]=kmatnew[key.Item1*3+key.Item2,key.Item1*3+key.Item2]+deltastif;
                kmatnew[dofindex,dofindex]=kmatnew[dofindex,dofindex]+deltastif;
                kmatnew[key.Item1*3+key.Item2,dofindex]=kmatnew[key.Item1*3+key.Item2,dofindex]-deltastif;
                kmatnew[dofindex,key.Item1*3+key.Item2]=kmatnew[dofindex,key.Item1*3+key.Item2]-deltastif;

            }
            return kmatnew;
        }
        public static List<Vector<double>> residualsforcereversestiffness(Vector<double> uhat,Vector<double> fhat,Dictionary<(int,int),List<double[]>>tnsprings,List<int> springmap)
        {
            var V=Vector<double>.Build;
            Vector<double> ghat=V.Dense(uhat.Count);
            Vector<double> freversehat=V.Dense(uhat.Count);
            Vector<double> stiffhold=V.Dense(uhat.Count);
            int dofindex;
            double[] Finthold;
            double u;
            double Fext;
            foreach((int,int) key in tnsprings.Keys)
            {
                dofindex=springmap[key.Item1*3+key.Item2];
                u=uhat[key.Item1*3+key.Item2];
                Fext=-fhat[dofindex];
                Finthold=nonlinearfunctions.Fintread(tnsprings[key],u);
                ghat[key.Item1*3+key.Item2]=Fext-Finthold[0];
                freversehat[dofindex]=(Fext-Finthold[0]);
                stiffhold[dofindex]=Finthold[1];
            }

            return new List<Vector<double>> {ghat,freversehat,stiffhold};
        }
        public static double[] Fintread(List<double[]> forcedisp, double u)
        {
            if (u <= forcedisp[0][0])
            {
                return new double[] {forcedisp[0][1]+(forcedisp[1][1]-forcedisp[0][1])*(u-forcedisp[0][0])/(forcedisp[1][0]-forcedisp[0][0]),(forcedisp[1][1]-forcedisp[0][1])/(forcedisp[1][0]-forcedisp[0][0])};
            }
            else if(u >= forcedisp[forcedisp.Count - 1][0])
            {
                return new double[] {forcedisp[forcedisp.Count - 2][1]+(forcedisp[forcedisp.Count - 1][1]-forcedisp[forcedisp.Count - 2][1])*(u-forcedisp[forcedisp.Count - 2][0])/(forcedisp[forcedisp.Count - 1][0]-forcedisp[forcedisp.Count - 2][0]),(forcedisp[forcedisp.Count - 1][1]-forcedisp[forcedisp.Count - 2][1])/(forcedisp[forcedisp.Count - 1][0]-forcedisp[forcedisp.Count - 2][0])};
            }
            else
            {
                int iter=0;
                while (u > forcedisp[iter+1][0])
                {
                    iter++;
                }
                return new double[] {forcedisp[iter][1]+(forcedisp[iter+1][1]-forcedisp[iter][1])*(u-forcedisp[iter][0])/(forcedisp[iter+1][0]-forcedisp[iter][0]),(forcedisp[iter+1][1]-forcedisp[iter][1])/(forcedisp[iter+1][0]-forcedisp[iter][0])};
            }
        }
    }

    class graphingtablefunctions
    {
        public static object[,] nodecoords(object[,] nodes)
        {
           object[,] outarray = new object[nodes.GetLength(0),2];
           for (int i = 0; i < nodes.GetLength(0); i++)
            {
                outarray[i,0]=nodes[i,0];
                outarray[i,1]=nodes[i,1];
            }
           return outarray;
        }
        public static object[,] elementcoords(object[,] nodes,object[,] elements)
        {
           object[,] outarray = new object[elements.GetLength(0)*3,2];
           for (int i = 0; i < elements.GetLength(0); i++)
            {
                outarray[i*3,0]=nodes[(int)((double)elements[i,0]-1),0];
                outarray[i*3,1]=nodes[(int)((double)elements[i,0]-1),1];
                outarray[i*3+1,0]=nodes[(int)((double)elements[i,1]-1),0];
                outarray[i*3+1,1]=nodes[(int)((double)elements[i,1]-1),1];
                outarray[i*3+2,0]=ExcelError.ExcelErrorNA;
                outarray[i*3+2,1]=ExcelError.ExcelErrorNA;
            }
           return outarray;
        }

        public static object[,] grapheffects(object[,] nodes,object[,] elements,object[,] results,object[,] requests)
        {
            requests[0,5]=requests[0,1];
            requests[0,6]=requests[0,2];
            requests[0,7]=requests[0,0];
            requests[0,8]="All";
            string dirn=requests[0,3].ToString() ?? string.Empty;
            int dirnindex;
            switch (dirn)
            {
                case "x":
                    dirnindex=0;
                    break;
                case "y":
                    dirnindex=1;
                    break;
                case "zz":
                    dirnindex=2;
                    break;
                default:
                    dirnindex=0;
                    break;
            }
            double scale=(double)requests[0,4];
            double xmin=(double)nodes[0,0];
            double xmax=(double)nodes[0,0];
            double ymin=(double)nodes[0,1];
            double ymax=(double)nodes[0,1];
            for (int i = 1; i < nodes.GetLength(0); i++)
            {
                xmin=Math.Min(xmin,(double)nodes[i,0]);
                xmax=Math.Max(xmax,(double)nodes[i,0]);
                ymin=Math.Min(ymin,(double)nodes[i,1]);
                ymax=Math.Max(ymax,(double)nodes[i,1]);
            }
            double graphrange=Math.Sqrt(Math.Pow(xmax-xmin,2)+Math.Pow(ymax-ymin,2))/2;
            int actionsdisplacements;
            int nodeselements;
            object[,] tabholder=tabularise(results,requests);
            if ((String)requests[0, 5] == "Actions")
            {
                actionsdisplacements=0;
            }
            else
            {
                actionsdisplacements=1;
            }
            if ((String)requests[0, 6] == "Nodes")
            {
                nodeselements=0;
            }
            else
            {
                nodeselements=2;
            }
            object[,] outholder=new object[0,0];
            switch (actionsdisplacements + nodeselements)
            {
                case 0:
                    outholder=graphingtablefunctions.graphnodeactions(nodes, tabholder,dirnindex,scale,graphrange);
                    break;
                case 1:
                    outholder=graphingtablefunctions.graphnodedisplacements(nodes, tabholder,scale);
                    break;
                case 2:
                    outholder=graphingtablefunctions.graphelementactions(nodes,elements, tabholder,dirnindex,scale,graphrange);
                    break;
                case 3:
                    outholder=graphingtablefunctions.graphelementdisplacements(nodes,elements, tabholder,scale);
                    break;
            }
            return outholder;

        }
        public static object[,] graphnodeactions(object[,] nodes, object[,] tabularised,int dirnindex,double scale,double graphrange)
        {
            List<object[]> outholder= new List<object[]>();
            double maxmag=0;
            double nodex;
            double nodey;
            double netscale;
            for (int i = 0; i < nodes.GetLength(0); i++)
            {
                maxmag=Math.Max(Math.Abs((double)tabularised[1+i,2+dirnindex]),maxmag);
            }
            if (maxmag == 0)
            {
                maxmag=1;
            }
            netscale=scale/100*graphrange/maxmag/2;
            for (int i = 0; i < nodes.GetLength(0); i++)
            {
                nodex=(double)nodes[i,0];
                nodey=(double)nodes[i,1];
                outholder.Add(new object[] {nodex,nodey});
                switch (dirnindex)
                {
                    case 0:
                        outholder.Add(new object[] {nodex+netscale*(double)tabularised[i+1,2+dirnindex],nodey});
                        break;
                    case 1:
                        outholder.Add(new object[] {nodex,nodey+netscale*(double)tabularised[i+1,2+dirnindex]});
                        break;
                    case 2:
                        outholder.Add(new object[] {nodex+netscale*((double)tabularised[i+1,2+dirnindex])/Math.Sqrt(2),nodey+netscale*((double)tabularised[i+1,2+dirnindex])/Math.Sqrt(2)});
                        break;
                }
                outholder.Add(new object[] {ExcelError.ExcelErrorNA,ExcelError.ExcelErrorNA});
            }
            object[,] outarray= new object[outholder.Count(), 2];
            for (int i=0; i<outholder.Count();i++)
            {
                outarray[i,0]=outholder[i][0];
                outarray[i,1]=outholder[i][1];
            }
            return outarray;
        }
        public static object[,] graphnodedisplacements(object[,] nodes, object[,] tabularised,double scale)
        {
            List<object[]> outholder= new List<object[]>();
            double nodex;
            double nodey;
            double netscale;
            
            netscale=scale;
            for (int i = 0; i < nodes.GetLength(0); i++)
            {
                nodex=(double)nodes[i,0];
                nodey=(double)nodes[i,1];
                outholder.Add(new object[] {nodex,nodey});
                outholder.Add(new object[] {nodex+netscale*((double)tabularised[i+1,2]),nodey+netscale*((double)tabularised[i+1,3])});
                outholder.Add(new object[] {ExcelError.ExcelErrorNA,ExcelError.ExcelErrorNA});
            }
            object[,] outarray= new object[outholder.Count(), 2];
            for (int i=0; i<outholder.Count();i++)
            {
                outarray[i,0]=outholder[i][0];
                outarray[i,1]=outholder[i][1];
            }
            return outarray;
        }
        public static object[,] graphelementactions(object[,] nodes, object[,] elements,object[,] tabularised,int dirnindex,double scale,double graphrange)
        {
            List<object[]> outholder= new List<object[]>();
            double maxmag=0;
            double elementx1;
            double elementx2;
            double elementy1;
            double elementy2;
            double xvec;
            double yvec;
            double magn;
            double effectmag;
            double netscale;
            double chapercent;
            for (int i = 0; i < nodes.GetLength(0); i++)
            {
                maxmag=Math.Max(Math.Abs((double)tabularised[1+i,4+dirnindex]),maxmag);
            }
            if (maxmag == 0)
            {
                maxmag=1;
            }
            netscale=scale/100*graphrange/maxmag/2;
            int tabiter=1;
            for (int i = 0; i < elements.GetLength(0); i++)
            {
                elementx1=(double)nodes[(int)((double)elements[i,0]-1),0];
                elementy1=(double)nodes[(int)((double)elements[i,0]-1),1];
                elementx2=(double)nodes[(int)((double)elements[i,1]-1),0];
                elementy2=(double)nodes[(int)((double)elements[i,1]-1),1];
                magn=Math.Sqrt(Math.Pow(elementx2-elementx1,2)+Math.Pow(elementy2-elementy1,2));
                xvec=(elementx2-elementx1)/magn;
                yvec=(elementy2-elementy1)/magn;
                outholder.Add(new object[] {elementx1,elementy1});
                while ((int)tabularised[tabiter, 1] == i + 1)
                {
                    effectmag=(double)tabularised[tabiter,4+dirnindex];
                    if (dirnindex == 2)
                    {
                        effectmag=-effectmag;
                    }
                    chapercent=(double)tabularised[tabiter,2];
                    outholder.Add(new object[] {elementx1+xvec*magn*chapercent-yvec*netscale*effectmag,elementy1+yvec*magn*chapercent+xvec*netscale*effectmag});
                    tabiter=tabiter+1;
                    if (tabiter == tabularised.GetLength(0))
                    {
                        break;
                    }
                }
                outholder.Add(new object[] {elementx2,elementy2});
                outholder.Add(new object[] {ExcelError.ExcelErrorNA,ExcelError.ExcelErrorNA});
            }
            object[,] outarray= new object[outholder.Count(), 2];
            for (int i=0; i<outholder.Count();i++)
            {
                outarray[i,0]=outholder[i][0];
                outarray[i,1]=outholder[i][1];
            }
            return outarray;
        }
        public static object[,] graphelementdisplacements(object[,] nodes, object[,] elements,object[,] tabularised,double scale)
        {
            List<object[]> outholder= new List<object[]>();
            double elementx1;
            double elementx2;
            double elementy1;
            double elementy2;
            double xvec;
            double yvec;
            double magn;
            double effectmagx;
            double effectmagy;
            double netscale;
            double chapercent;
            
            netscale=scale;
            int tabiter=1;
            for (int i = 0; i < elements.GetLength(0); i++)
            {
                elementx1=(double)nodes[(int)((double)elements[i,0]-1),0];
                elementy1=(double)nodes[(int)((double)elements[i,0]-1),1];
                elementx2=(double)nodes[(int)((double)elements[i,1]-1),0];
                elementy2=(double)nodes[(int)((double)elements[i,1]-1),1];
                magn=Math.Sqrt(Math.Pow(elementx2-elementx1,2)+Math.Pow(elementy2-elementy1,2));
                xvec=(elementx2-elementx1)/magn;
                yvec=(elementy2-elementy1)/magn;
                while ((int)tabularised[tabiter, 1] == i + 1)
                {
                    effectmagx=(double)tabularised[tabiter,4]*xvec-(double)tabularised[tabiter,5]*yvec;
                    effectmagy=(double)tabularised[tabiter,4]*yvec+(double)tabularised[tabiter,5]*xvec;
                    chapercent=(double)tabularised[tabiter,2];
                    outholder.Add(new object[] {elementx1+xvec*magn*chapercent+netscale*effectmagx,elementy1+yvec*magn*chapercent+netscale*effectmagy});
                    tabiter=tabiter+1;
                    if (tabiter == tabularised.GetLength(0))
                    {
                        break;
                    }
                }
                outholder.Add(new object[] {ExcelError.ExcelErrorNA,ExcelError.ExcelErrorNA});
            }
            object[,] outarray= new object[outholder.Count(), 2];
            for (int i=0; i<outholder.Count();i++)
            {
                outarray[i,0]=outholder[i][0];
                outarray[i,1]=outholder[i][1];
            }
            return outarray;
        }
        public static object[,] tabularise(object[,] results,object[,] requests)
        {
            int actionsdisplacements;
            int nodeselements;
            string lcrequest;
            string memberrequest;   
            if ((String)requests[0, 5] == "Actions")
            {
                actionsdisplacements=0;
            }
            else
            {
                actionsdisplacements=1;
            }
            if ((String)requests[0, 6] == "Nodes")
            {
                nodeselements=0;
            }
            else
            {
                nodeselements=2;
            }
            if ((requests[0, 7].ToString()?? string.Empty).ToLower() == "all")
            {
                lcrequest="-1";
            }
            else
            {
                lcrequest=requests[0,7].ToString()?? string.Empty;
            }
            if ((requests[0, 8].ToString()?? string.Empty).ToLower() == "all")
            {
                memberrequest="-1";
            }
            else
            {
                memberrequest=requests[0,8].ToString() ?? string.Empty;
            }
            List<object[]> lcmemberstext=lcmembers(results,actionsdisplacements + nodeselements);
            SortedSet<int> lcall=new SortedSet<int>();
            SortedSet<int> memberall=new SortedSet<int>();
            for (int i = 0; i < lcmemberstext.Count; i++)
            {
                lcall.Add(int.Parse((string)lcmemberstext[i][0]));
                memberall.Add(int.Parse((string)lcmemberstext[i][1]));
            }
            List<int> lclist = stiffnessmatcalcs.lcmemberstringread(lcrequest, lcall.ElementAt(lcall.Count-1),lcall.ToList());
            List<int> memberlist = stiffnessmatcalcs.lcmemberstringread(memberrequest, memberall.ElementAt(memberall.Count-1),memberall.ToList());
            SortedSet<int> lcset=new SortedSet<int>(lclist);
            SortedSet<int> memberset=new SortedSet<int>(memberlist);
            for (int i = lcmemberstext.Count - 1; i >= 0; i--)
            {
                if (!(lcset.Contains(int.Parse((String)lcmemberstext[i][0]))) || !(memberset.Contains(int.Parse((String)lcmemberstext[i][1]))))
                {
                    lcmemberstext.RemoveAt(i);
                }
            }
            List<object[]> outlist= new List<object[]>();
            switch (actionsdisplacements + nodeselements)
            {
                case 0:
                    outlist.Add(new object[] {"Load Case", "Node Index","x Reaction (N)","y Reaction (N)","zz Reaction (Nm)"});
                    outlist.AddRange(noderead(lcmemberstext));
                    break;
                case 1:
                    outlist.Add(new object[] {"Load Case", "Node Index","x Displacement (m)","y Displacement (m)","zz rotation (rad)"});
                    outlist.AddRange(noderead(lcmemberstext));
                    break;
                case 2:
                    outlist.Add(new object[] {"Load Case", "Element Index","Element Cha (%)","Element Cha (m)","Axial Force (N)","Shear Force (N)","Bending Moment (Nm)"});
                    outlist.AddRange(elementread(lcmemberstext));
                    break;
                case 3:
                    outlist.Add(new object[] {"Load Case", "Element Index","Element Cha (%)","Element Cha (m)","Axial displacement (m)","Shear displacement (m))","Rotation (rad)"});
                    outlist.AddRange(elementread(lcmemberstext));
                    break;
            }
            

            object[,] outarray=new object[outlist.Count,outlist[0].GetLength(0)];
            for (int i=0; i < outlist.Count; i++)
            {
                for (int j=0; j < outlist[i].GetLength(0); j++)
                {
                    outarray[i,j]=outlist[i][j];
                }
            }
            return outarray;
        }
        public static List<object[]> lcmembers(object[,] results,int index)
        {
            string txtholder="";
            for (int i = 0; i < results.GetLength(1); i++)
            {
                txtholder=txtholder+results[index,i];
            }
            txtholder=txtholder.Split("$")[1];
            List<object[]> outlist=new List<object[]>();
            string[] LC=txtholder.Split(":");
            string[] member;
            object[] outrow;
            string[] lcandmemberandresults;
            for (int i = 0; i < LC.GetLength(0); i++)
            {
                member=LC[i].Split(";");
                for (int j = 0; j < member.GetLength(0); j++)
                {
                    lcandmemberandresults=member[j].Split("/");
                    outrow= new object[] {lcandmemberandresults[0].Split("~")[0],lcandmemberandresults[0].Split("~")[1],lcandmemberandresults[1]};
                    outlist.Add(outrow);
                }
            }
            return outlist;
        }
        public static List<object[]> noderead(List<object[]> lcmemberstext)
        {
            object[] rw;
            string[] txtsplittemp;
            List<object[]> outlist= new List<object[]>();
            for (int i = 0; i < lcmemberstext.Count; i++)
            {
                txtsplittemp=((String)lcmemberstext[i][2]).Split(",");
                rw= new object[] {int.Parse((String)lcmemberstext[i][0]),int.Parse((String)lcmemberstext[i][1]),double.Parse((String)txtsplittemp[0]),double.Parse((String)txtsplittemp[1]),double.Parse((String)txtsplittemp[2])};
                outlist.Add(rw);
            }
            return outlist;
        }
        public static List<object[]> elementread(List<object[]> lcmemberstext)
        {
            object[] rw;
            string[] txtsplittemp;
            string[] txtsplittemp2;
            string[] txtsplittemp3;
            string[] txtsplittemp4;

            List<object[]> outlist= new List<object[]>();
            for (int i = 0; i < lcmemberstext.Count; i++)
            {
                txtsplittemp=((String)lcmemberstext[i][2]).Split("+");
                for (int j = 0; j < txtsplittemp.Count(); j++)
                {
                    txtsplittemp2=(txtsplittemp[j]).Split("#");
                    txtsplittemp3=(txtsplittemp2[1]).Split(",");
                    txtsplittemp4=(txtsplittemp2[0]).Split("|");
                    rw= new object[] {int.Parse((String)lcmemberstext[i][0]),int.Parse((String)lcmemberstext[i][1]),double.Parse((String)txtsplittemp4[0]),double.Parse((String)txtsplittemp4[1]),double.Parse((String)txtsplittemp3[0]),double.Parse((String)txtsplittemp3[1]),double.Parse((String)txtsplittemp3[2])};
                    outlist.Add(rw);
                }
            }
            return outlist;
        }
    }
    class parseclass
    {
        public static double parseloadequations(double lc, double index, double x, double y,double zz,string eqn)
        {
             System.Data.DataTable table = new System.Data.DataTable ();

            // Create the first column.
            DataColumn LCColumn = new DataColumn();
            LCColumn.DataType = System.Type.GetType("System.Decimal");
            LCColumn.ColumnName = "lc";
            LCColumn.DefaultValue = lc;
            
            DataColumn indexColumn = new DataColumn();
            indexColumn.DataType = System.Type.GetType("System.Decimal");
            indexColumn.ColumnName = "index";
            indexColumn.DefaultValue = index;

            DataColumn xColumn = new DataColumn();
            xColumn.DataType = System.Type.GetType("System.Decimal");
            xColumn.ColumnName = "x";
            xColumn.DefaultValue = x;

            DataColumn yColumn = new DataColumn();
            yColumn.DataType = System.Type.GetType("System.Decimal");
            yColumn.ColumnName = "y";
            yColumn.DefaultValue = y;

            DataColumn zzColumn = new DataColumn();
            zzColumn.DataType = System.Type.GetType("System.Decimal");
            zzColumn.ColumnName = "zz";
            zzColumn.DefaultValue = zz;;

            DataColumn loadColumn = new DataColumn();
            loadColumn.DataType = System.Type.GetType("System.Decimal");
            loadColumn.ColumnName = "load";
            loadColumn.Expression = eqn;



            // Add columns to DataTable.
            table.Columns.Add(LCColumn);
            table.Columns.Add(indexColumn);
            table.Columns.Add(xColumn);
            table.Columns.Add(yColumn);
            table.Columns.Add(zzColumn);
            table.Columns.Add(loadColumn);

            DataRow row = table.NewRow();
            table.Rows.Add(row);

            double ld = double.Parse(table.Rows[0]["load"].ToString() ?? "0");
            return ld;
        }
    }
    class controllerclass
    {
        public static string[,] controller(double[,] tnodes,double[,] telements,double[,] tloads,double[,] tbloads,object[,] extracts, Dictionary<(int,int),List<double[]>> tnsprings)
        {
            var M=Matrix<double>.Build;
            double[,] beamgeom=stiffnessmatcalcs.beamgeom(tnodes,telements);
            HashSet<int> lclist=stiffnessmatcalcs.LClist(tloads,tbloads);
            List<int> springmap=stiffnessmatcalcs.springmap(tnodes);
            Matrix<double> kmat=stiffnessmatcalcs.kmat(tnodes,telements,beamgeom);
            Matrix<double> tmat=stiffnessmatcalcs.tmat(tnodes,telements,beamgeom);
            Matrix<double> kmattransf = (((tmat.Transpose()).Multiply(kmat)).Multiply(tmat));
            int nodecount=tnodes.GetLength(0);
            int elementcount=telements.GetLength(0);
            int springcount=stiffnessmatcalcs.springcount(tnodes);
            double[] filtermap=stiffnessmatcalcs.filtermap(tnodes,springcount);
            Matrix<double> kmatreduce=stiffnessmatcalcs.kmatreduce(kmattransf,filtermap);
            Matrix<double> invkmatreduce=kmatreduce.Inverse();
            Dictionary<int,Dictionary<int,List<int>>> nodeloadsholder=stiffnessmatcalcs.nodeloadsholder(tloads);
            Dictionary<int,Dictionary<int,List<double[,]>>> beamloadsholder=stiffnessmatcalcs.beamloadsholder(tbloads,beamgeom);
            Dictionary<int,Dictionary<int,SortedSet<double>>> beamchapoints=stiffnessmatcalcs.beamchapoints(beamloadsholder);
            Vector<double> fvector;
            Vector<double> Ttfvector;
            Vector<double> mapttfvector;
            Vector<double> uhatvectorred;
            Vector<double> uhatvector;
            Vector<double> fhatvector;
            Vector<double> u;
            Vector<double> f;
            Matrix<double> results;
            Dictionary<int,List<Vector<double>>> fextholder=new Dictionary<int,List<Vector<double>>>();
            Dictionary<int,Matrix<double>> resultsholder=new Dictionary<int,Matrix<double>>();
            Dictionary<int,Matrix<double>> resultsholder2=new Dictionary<int,Matrix<double>>();
            Dictionary<int,List<int>> nodein;
            Dictionary<int,List<double[,]>> beamin;
            foreach(var key in lclist)
            {
                if (nodeloadsholder.ContainsKey(key))
                {
                    nodein=nodeloadsholder[key];
                }
                else
                {
                    nodein=new Dictionary<int,List<int>>();
                }
                 if (beamloadsholder.ContainsKey(key))
                {
                    beamin=beamloadsholder[key];
                }
                else
                {
                    beamin=new Dictionary<int,List<double[,]>>();
                }
                fvector=stiffnessmatcalcs.Fvector(tloads,kmat.RowCount,nodein,beamgeom,telements,beamin);
                Ttfvector=(tmat.Transpose()).Multiply(fvector);
                fextholder.Add(key,new List<Vector<double>>{fvector,Ttfvector});
                mapttfvector=stiffnessmatcalcs.filterby(Ttfvector,filtermap);
                uhatvectorred=invkmatreduce.Multiply(mapttfvector);
                uhatvector=stiffnessmatcalcs.unfilterby(uhatvectorred,filtermap,Ttfvector.Count);
                fhatvector=kmattransf.Multiply(uhatvector);
                if (tnsprings.Count>0)
                {
                    List<Vector<double>> nonlin=nonlinearfunctions.nonlinspring(uhatvector,fhatvector,kmattransf,tnsprings,springmap,tnodes,filtermap);
                    uhatvector=nonlin[0];
                    fhatvector=nonlin[1];
                }
                u=tmat.Multiply(uhatvector);
                f=kmat.Multiply(u);
                results=Matrix<double>.Build.DenseOfColumnVectors(u,f);
                resultsholder.Add(key,results);
                results=Matrix<double>.Build.DenseOfColumnVectors(uhatvector,fhatvector);
                resultsholder2.Add(key,results);
            }
            //Matrix<double> outmatrix=M.Dense(kmat.RowCount,0);
            //foreach(var key in resultsholder.Keys)
            //{
                //outmatrix=outmatrix.Append(resultsholder[key]);
            //}
            string[,] extracted=stiffnessmatcalcs.extraction(extracts,resultsholder,resultsholder2,nodeloadsholder,fextholder,tloads,nodecount,elementcount,springmap,beamgeom,beamloadsholder,beamchapoints);

            

            return extracted;
        }
    }
}