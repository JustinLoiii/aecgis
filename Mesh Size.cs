using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.DataSourcesFile;
using ESRI.ArcGIS.Geodatabase;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ESRI.ArcGIS.Geoprocessor;
using ESRI.ArcGIS.AnalysisTools;
using ESRI.ArcGIS.DataManagementTools;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Display;
using System.Reflection;
using System.Data.OleDb;
using ADOX;
using System.Collections;
using ESRI.ArcGIS.ADF;


/*
做了2件事
1、加载研究区shapefile文件（先判断坐标系为unknow，再判断是否有投影，如果满足条件就加载），在加载的时候复制一份到Others文件夹里
2、创建渔网并赋值，并将新生成的网格group layer
*/

namespace NEW_FISH
{
    public partial class 渔网Form : Form
    {
        public 渔网Form()
        {
            InitializeComponent();
        }

        IFeatureClass pFeatureClass;
        IFeatureLayer pFeatureLyr;
        IWorkspaceFactory pWorkspaceFactory;
        IFeatureWorkspace pFeaWorkspace;
        string filepath;//选择的研究区范围
        string purename;//不包含后缀的名称
        public static string text;
        public static string workpath;
        public static string txtpath;
        public static string mxdPath;
        public static string mdbPath;
        public static string othersPath;
        public static string gridPath;
        public static string databasePath;
        private void button1_Click(object sender, EventArgs e)
        {
            //获取当前路径
            mxdPath = GetActiveDocumentPath(ArcMap.Document as IMapDocument);
            mdbPath= mxdPath.Replace(".mxd", ".mdb");
            othersPath = mxdPath.Replace(".mxd", "DATABASE")+"\\"+ "Others";
            gridPath= mxdPath.Replace(".mxd", "DATABASE") + "\\" + "GRID";
            databasePath = mxdPath.Replace(".mxd", "DATABASE");
            IMxDocument mxd = ArcMap.Document;
            IActiveView act = mxd.ActiveView;
            IMap mypmap = act.FocusMap;

                //读取选取的shapefile文件
                OpenFileDialog file = new OpenFileDialog();
                file.Filter = "ShapeFile文件|*.shp";//限制后缀名
                if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    //变量定义
                    filepath = file.FileName;//获取包括文件名的总目录
                    string path = System.IO.Path.GetDirectoryName(filepath);//获取父目录
                    string allname = System.IO.Path.GetFileName(filepath);//获取包含后缀名的文件名
                    string[] files = Directory.GetFiles(path);//获取父目录中所有文件
                    purename = System.IO.Path.GetFileNameWithoutExtension(filepath);
                    string strFileName = path + "\\" + purename + ".prj";
                    //判断坐标系是否为unknown
                    if (!System.IO.File.Exists(strFileName))
                    {
                        MessageBox.Show("该ShapeFile文件没有坐标系统！请重新输入", "提示", MessageBoxButtons.OK);
                        textBox1.Clear();
                        textBox1.Enabled = false;
                    }

                    else if (System.IO.File.Exists(strFileName))
                    {
                        //判断是否有投影坐标
                        //预加载shapefile
                        pWorkspaceFactory = new ShapefileWorkspaceFactory();
                        pFeaWorkspace = pWorkspaceFactory.OpenFromFile(path, 0) as IFeatureWorkspace;
                        pFeatureClass = pFeaWorkspace.OpenFeatureClass(allname);  //打开文件
                        pFeatureLyr = new FeatureLayer();
                        pFeatureLyr.FeatureClass = pFeatureClass;
                        pFeatureLyr.Name = pFeatureLyr.FeatureClass.AliasName;
                        //判断是否有投影
                        text = System.IO.File.ReadAllText(strFileName);
                        int start = 0, length = 6;
                        string front = text.Substring(start, length);
                        if (front == "GEOGCS")
                        {
                            MessageBox.Show("该shapefile没有投影坐标系！", "提示", MessageBoxButtons.OK);
                            textBox1.Enabled = false;
                        }
                        //有投影坐标系，加载
                        else if (front == "PROJCS")
                        {
                            //string linepath = arr[2].ToString();


                            mypmap.AddLayer(pFeatureLyr);
                            ESRI.ArcGIS.Geoprocessor.Geoprocessor GP = new ESRI.ArcGIS.Geoprocessor.Geoprocessor();
                            GP.OverwriteOutput = true;
                            ESRI.ArcGIS.DataManagementTools.Copy copyShp = new ESRI.ArcGIS.DataManagementTools.Copy();
                            copyShp.in_data = filepath;
                            copyShp.out_data = othersPath + "\\" + purename + "_copy.shp";
                            GP.AddOutputsToMap = false;
                            GP.Execute(copyShp, null);
                            textBox1.Text = allname;
                            textBox1.Enabled = true;


                            //将复制的研究区范围添加进access数据库
                            //需要判断是不是第二次输入，然后把第一次access中的内容删除再添加
                            string connStr = @"Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:DataBase Password=;Data Source=" + mdbPath;
                            OleDbConnection tempconn = new OleDbConnection(connStr);
                            tempconn.Open();
                            //将输入图层的坐标存入access数据库
                            //需要判断是不是第二次输入，然后把第一次access中的内容删除再添加
                            OleDbCommand cmd1 = new OleDbCommand("delete * From [Others] where [Coord] = '研究区坐标'", tempconn);
                            cmd1.ExecuteNonQuery();

                            string strIter1 = "Alter TABLE Others Alter COLUMN ID COUNTER (1, 1)";
                            OleDbCommand oleDbCommand4 = new OleDbCommand(strIter1, tempconn);
                            oleDbCommand4.ExecuteNonQuery();

                            string sql2 = "INSERT INTO [Others] ([Coord],[Text]) values('" + "研究区坐标" + "', '" + text + "')";
                            OleDbCommand oleDbCommand2 = new OleDbCommand(sql2, tempconn);
                            oleDbCommand2.ExecuteNonQuery();
                            tempconn.Close();
                            mxd.UpdateContents();
                            act.Refresh();
                            MessageBox.Show("ShapeFile文件加载成功！", "提示", MessageBoxButtons.OK);

                        }
                        else
                        {
                            MessageBox.Show("发生未知错误！", "提示", MessageBoxButtons.OK);
                        }
                    }
                
            }                  
        }

        public static int nlayers;
        public static int nrows;
        public static int ncolums;
        public static double rextent;
        public static double cextent;
        public static IGroupLayer groupLayer;
        private void button4_Click(object sender, EventArgs e)
        {
            mxdPath = GetActiveDocumentPath(ArcMap.Document as IMapDocument);
            mdbPath = mxdPath.Replace(".mxd", ".mdb");
            othersPath = mxdPath.Replace(".mxd", "DATABASE") + "\\" + "Others";
            gridPath = mxdPath.Replace(".mxd", "DATABASE") + "\\" + "GRID";
            databasePath = mxdPath.Replace(".mxd", "DATABASE");


            this.Refresh();
            IMxDocument mxd = ArcMap.Document;
            IActiveView act = mxd.ActiveView;
            IMap mypmap = act.FocusMap;
            nlayers = Convert.ToInt32(NLayers.Text);
            rextent = Convert.ToDouble(RExtent.Text);
            cextent = Convert.ToDouble(CExtent.Text);
            nrows = Convert.ToInt32(NRows.Text);
            ncolums= Convert.ToInt32(NColums.Text);
            double coordx=Convert.ToDouble(CoordX.Text);
            double coordy = Convert.ToDouble(CoordY.Text);
            double yaxisx= Convert.ToDouble(textBox2.Text);
            groupLayer = new GroupLayerClass();
            groupLayer.Name = "含水层";
            act.Refresh();
            mxd.UpdateContents();
            try
            {
                for (int i =0;i<nlayers;i++)
                {
                    //创建渔网
                    Geoprocessor gp = new Geoprocessor();
                    gp.OverwriteOutput = true;
                    gp.SetEnvironmentValue("outputCoordinateSystem",text);
                    //gp.SetEnvironmentValue("CartographicCoordinateSystem",text); 
                    ESRI.ArcGIS.DataManagementTools.CreateFishnet fish = new ESRI.ArcGIS.DataManagementTools.CreateFishnet();
                    fish.out_feature_class = gridPath + "\\" + "layer"+i+".shp";//输出路径需修改
                    fish.origin_coord = coordx.ToString() + " " + coordy.ToString();
                    fish.y_axis_coord = (coordx+(Math.Tan(yaxisx/180*Math.PI) * (cextent * nrows))).ToString() + " " + (coordy+(cextent * nrows)).ToString();
                    fish.cell_width = rextent;
                    fish.cell_height = cextent;
                    fish.number_rows = nrows;
                    fish.number_columns = ncolums;
                    fish.geometry_type = "POLYGON";
                    fish.labels= "NO_LABELS";                    
                    gp.AddOutputsToMap = false;
                    gp.Execute(fish, null);
                    

                    //给属性表中每一个网格赋予行列号
                    pWorkspaceFactory = new ShapefileWorkspaceFactory();
                    pFeaWorkspace = pWorkspaceFactory.OpenFromFile(gridPath, 0) as IFeatureWorkspace;//工作空间路径需修改
                    pFeatureClass = pFeaWorkspace.OpenFeatureClass("layer" + i + ".shp");  //打开文件
                    IField pField = new FieldClass();
                    IFieldEdit pFieldEdit = pField as IFieldEdit;
                    pFieldEdit.Name_2 = "Rows";
                    pFieldEdit.AliasName_2 = "Rows";
                    pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble;
                    pFieldEdit.Length_2 = 20;
                    pFeatureClass.AddField(pField);
                    pFieldEdit.Name_2 = "Colums";
                    pFieldEdit.AliasName_2 = "Colums";
                    pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble;
                    pFieldEdit.Length_2 = 20;
                    pFeatureClass.AddField(pField);
                    pFieldEdit.Name_2 = "行号";
                    pFieldEdit.AliasName_2 = "行号";
                    pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble;
                    pFieldEdit.Length_2 = 20;
                    pFeatureClass.AddField(pField);
                    pFieldEdit.Name_2 = "列号";
                    pFieldEdit.AliasName_2 = "列号";
                    pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble;
                    pFieldEdit.Length_2 = 20;
                    pFeatureClass.AddField(pField);
                    pFieldEdit.Name_2 = "IBOUND";
                    pFieldEdit.AliasName_2 = "IBOUND";
                    pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger;
                    pFieldEdit.Length_2 = 20;
                    pFeatureClass.AddField(pField);
                    pFieldEdit.Name_2 = "最小X";
                    pFieldEdit.AliasName_2 = "最小X";
                    pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble;
                    pFieldEdit.Length_2 = 20;
                    pFeatureClass.AddField(pField);
                    pFieldEdit.Name_2 = "最大X";
                    pFieldEdit.AliasName_2 = "最大X";
                    pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble;
                    pFieldEdit.Length_2 = 20;
                    pFeatureClass.AddField(pField);
                    pFieldEdit.Name_2 = "最小Y";
                    pFieldEdit.AliasName_2 = "最小Y";
                    pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble;
                    pFieldEdit.Length_2 = 20;
                    pFeatureClass.AddField(pField);
                    pFieldEdit.Name_2 = "最大Y";
                    pFieldEdit.AliasName_2 = "最大Y";
                    pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble;
                    pFieldEdit.Length_2 = 20;
                    pFeatureClass.AddField(pField);
                    pFieldEdit.Name_2 = "宽度";
                    pFieldEdit.AliasName_2 = "宽度";
                    pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble;
                    pFieldEdit.Length_2 = 20;
                    pFeatureClass.AddField(pField);
                    pFieldEdit.Name_2 = "高度";
                    pFieldEdit.AliasName_2 = "高度";
                    pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble;
                    pFieldEdit.Length_2 = 20;
                    pFeatureClass.AddField(pField);
                    pFieldEdit.Name_2 = "顶板";
                    pFieldEdit.AliasName_2 = "顶板";
                    pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble;
                    pFieldEdit.Length_2 = 20;
                    pFeatureClass.AddField(pField);
                    pFieldEdit.Name_2 = "底板";
                    pFieldEdit.AliasName_2 = "底板";
                    pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble;
                    pFieldEdit.Length_2 = 20;
                    pFeatureClass.AddField(pField);


                    //给行号和列号赋值
                    ESRI.ArcGIS.DataManagementTools.CalculateField grid = new ESRI.ArcGIS.DataManagementTools.CalculateField();
                    grid.in_table = fish.out_feature_class;
                    grid.field = "Rows";
                    grid.expression = nrows;
                    gp.Execute(grid, null);
                    grid = null;
                    ESRI.ArcGIS.DataManagementTools.CalculateField grid1 = new ESRI.ArcGIS.DataManagementTools.CalculateField();
                    grid1.in_table = fish.out_feature_class;
                    grid1.field = "Colums";
                    grid1.expression = ncolums;
                    gp.Execute(grid1, null);
                    ESRI.ArcGIS.DataManagementTools.CalculateField grid2 = new ESRI.ArcGIS.DataManagementTools.CalculateField();
                    grid2.in_table = grid1.in_table;
                    grid2.field = "行号";
                    grid2.expression = "Abs ((Fix (( [FID] -1 +1)/ [Colums] )+1)- [Rows] )";
                    gp.Execute(grid2, null);
                    ESRI.ArcGIS.DataManagementTools.CalculateField grid3 = new ESRI.ArcGIS.DataManagementTools.CalculateField();
                    grid3.in_table = grid2.in_table;
                    grid3.field = "列号";
                    grid3.expression = "(( [FID] + [Colums] )mod [Colums] )";
                    gp.Execute(grid3, null);
                    ESRI.ArcGIS.DataManagementTools.CalculateField grid4 = new ESRI.ArcGIS.DataManagementTools.CalculateField();
                    grid4.in_table = grid3.in_table;
                    grid4.field = "Id";
                    grid4.expression = i +"* [Rows] * [Colums] + [行号] * [Colums] + [列号]";
                    gp.Execute(grid4, null);
                    ESRI.ArcGIS.DataManagementTools.DeleteField grid5 = new ESRI.ArcGIS.DataManagementTools.DeleteField();
                    grid5.in_table = grid4.in_table;
                    grid5.drop_field = "Rows;Colums";
                    gp.Execute(grid5, null);
                    ESRI.ArcGIS.DataManagementTools.CalculateField grid6 = new ESRI.ArcGIS.DataManagementTools.CalculateField();
                    grid6.in_table = grid5.in_table;
                    grid6.field = "最小X";
                    grid6.expression_type = "PYTHON_9.3";
                    grid6.expression = "!shape.extent.XMin!";
                    gp.Execute(grid6, null);
                    ESRI.ArcGIS.DataManagementTools.CalculateField grid7 = new ESRI.ArcGIS.DataManagementTools.CalculateField();
                    grid7.in_table = grid6.in_table;
                    grid7.field = "最大X";
                    grid7.expression_type = "PYTHON_9.3";
                    grid7.expression = "!shape.extent.XMax!";
                    gp.Execute(grid7, null);
                    ESRI.ArcGIS.DataManagementTools.CalculateField grid8 = new ESRI.ArcGIS.DataManagementTools.CalculateField();
                    grid8.in_table = grid7.in_table;
                    grid8.field = "宽度";
                    grid8.expression_type = "PYTHON_9.3";
                    grid8.expression = "!最大x! - !最小x!";
                    gp.Execute(grid8, null);
                    ESRI.ArcGIS.DataManagementTools.CalculateField grid9 = new ESRI.ArcGIS.DataManagementTools.CalculateField();
                    grid9.in_table = grid8.in_table;
                    grid9.field = "最小Y";
                    grid9.expression_type = "PYTHON_9.3";
                    grid9.expression = "!shape.extent.YMin!";
                    gp.Execute(grid9, null);
                    ESRI.ArcGIS.DataManagementTools.CalculateField grid10 = new ESRI.ArcGIS.DataManagementTools.CalculateField();
                    grid10.in_table = grid9.in_table;
                    grid10.field = "最大Y";
                    grid10.expression_type = "PYTHON_9.3";
                    grid10.expression = "!shape.extent.YMax!";
                    gp.Execute(grid10, null);
                    ESRI.ArcGIS.DataManagementTools.CalculateField grid11 = new ESRI.ArcGIS.DataManagementTools.CalculateField();
                    grid11.in_table = grid10.in_table;
                    grid11.field = "高度";
                    grid11.expression_type = "PYTHON_9.3";
                    grid11.expression = "!最大Y! - !最小Y!";
                    gp.Execute(grid11, null);
                    ESRI.ArcGIS.DataManagementTools.DeleteField grid12 = new ESRI.ArcGIS.DataManagementTools.DeleteField();
                    grid12.in_table = grid11.in_table;
                    grid12.drop_field = "最小X;最大X;最小Y;最大Y";
                    gp.Execute(grid12, null);





                    //把workspace文件夹中的N层网格加入到grouplayer：含水层中
                    pFeaWorkspace = pWorkspaceFactory.OpenFromFile(gridPath, 0) as IFeatureWorkspace;
                    pFeatureClass = pFeaWorkspace.OpenFeatureClass("layer" + i + ".shp");  //打开文件
                    pFeatureLyr = new FeatureLayer();
                    pFeatureLyr.FeatureClass = pFeatureClass;
                    pFeatureLyr.Name = pFeatureLyr.FeatureClass.AliasName;
                    groupLayer.Add(pFeatureLyr);


                }
                mypmap.AddLayer(groupLayer);
                act.Refresh();

                //新建access行宽表
                ADOX.Catalog catalog = new Catalog();
                ADODB.Connection cn = new ADODB.Connection();
                cn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdbPath, null, null, -1);
                catalog.ActiveConnection = cn;
                ADOX.Table table = new ADOX.Table();
                table.Name = "行间距";
                ADOX.Column column = new ADOX.Column();
                column.ParentCatalog = catalog;
                column.Name = "ID";
                column.Type = DataTypeEnum.adInteger;
                column.DefinedSize = 9;
                column.Properties["AutoIncrement"].Value = true;
                table.Columns.Append(column, DataTypeEnum.adInteger, 9);
                table.Keys.Append("FirstTablePrimaryKey", KeyTypeEnum.adKeyPrimary, column, null, null);
                table.Columns.Append("行号", DataTypeEnum.adVarWChar, 150);
                table.Columns.Append("行间距", DataTypeEnum.adVarWChar, 150);
                catalog.Tables.Append(table);
                //新建access列宽表
                ADOX.Table table1 = new ADOX.Table();
                ADOX.Column column1 = new ADOX.Column();
                table1.Name = "列间距";
                column1.ParentCatalog = catalog;
                column1.Name = "ID";
                column1.Type = DataTypeEnum.adInteger;
                column1.DefinedSize = 9;
                column1.Properties["AutoIncrement"].Value = true;
                table1.Columns.Append(column1, DataTypeEnum.adInteger, 9);
                table1.Keys.Append("FirstTablePrimaryKey", KeyTypeEnum.adKeyPrimary, column1, null, null);
                table1.Columns.Append("列号", DataTypeEnum.adVarWChar, 150);
                table1.Columns.Append("列间距", DataTypeEnum.adVarWChar, 150);
                catalog.Tables.Append(table1);

                string Rowstr = @"Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:DataBase Password=;Data Source=" + mdbPath;
                OleDbConnection Rowtempconn = new OleDbConnection(Rowstr);
                OleDbCommand Rowcmd = new OleDbCommand("delete * From 行间距", Rowtempconn);
                Rowtempconn.Open();
                Rowcmd.ExecuteNonQuery();

                string Rowstr1 = "Alter TABLE [行间距] Alter COLUMN ID COUNTER (1, 1)";
                OleDbCommand Rowcmd1 = new OleDbCommand(Rowstr1, Rowtempconn);
                Rowcmd1.ExecuteNonQuery();

                pFeaWorkspace = pWorkspaceFactory.OpenFromFile(gridPath, 0) as IFeatureWorkspace;
                pFeatureClass = pFeaWorkspace.OpenFeatureClass("layer0.shp");  //打开文件
                System.Data.DataTable dataTable1 = GetAttributesTable(pFeatureClass);
                DeleteSameRow(dataTable1, "行号");
                var lista6 = dataTable1.AsEnumerable().Select(d => d.Field<string>("行号")).ToList();
                var listb6 = dataTable1.AsEnumerable().Select(d => d.Field<string>("高度")).ToList();
                for (int i = 0; i < nrows; i++)
                {
                    string layer1 = "INSERT INTO [行间距] ([行号],[行间距]) values('" + lista6[i] + "', '" + listb6[i]  + "')";
                    OleDbCommand oleDbCommand = new OleDbCommand(layer1, Rowtempconn);
                    oleDbCommand.ExecuteNonQuery();
                }
                Rowtempconn.Close();






                string Colstr = @"Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:DataBase Password=;Data Source=" + mdbPath;
                OleDbConnection Coltempconn = new OleDbConnection(Colstr);
                OleDbCommand Colcmd = new OleDbCommand("delete * From 列间距", Coltempconn);
                Coltempconn.Open();
                Colcmd.ExecuteNonQuery();

                string Colstr1 = "Alter TABLE [列间距] Alter COLUMN ID COUNTER (1, 1)";
                OleDbCommand Colcmd1 = new OleDbCommand(Colstr1, Coltempconn);
                Colcmd1.ExecuteNonQuery();

                pFeaWorkspace = pWorkspaceFactory.OpenFromFile(gridPath, 0) as IFeatureWorkspace;
                pFeatureClass = pFeaWorkspace.OpenFeatureClass("layer0.shp");  //打开文件
                System.Data.DataTable dataTable2 = GetAttributesTable(pFeatureClass);
                DeleteSameRow(dataTable2, "列号");
                var lista7 = dataTable2.AsEnumerable().Select(d => d.Field<string>("列号")).ToList();
                var listb7 = dataTable2.AsEnumerable().Select(d => d.Field<string>("宽度")).ToList();
                for (int i = 0; i < ncolums; i++)
                {
                    string layer1 = "INSERT INTO [列间距] ([列号],[列间距]) values('" + lista7[i] + "', '" + listb7[i] + "')";
                    OleDbCommand oleDbCommand1 = new OleDbCommand(layer1, Coltempconn);
                    oleDbCommand1.ExecuteNonQuery();
                }
                Coltempconn.Close();


                pFeatureClass = pFeaWorkspace.OpenFeatureClass("layer1.shp");  //打开文件
                System.Data.DataTable dataTable3 = GetAttributesTable(pFeatureClass);
                string a5 = "FID";
                string b5 = "宽度";
                string c5 = "高度";
                System.Data.DataTable newTable5 = dataTable3.DefaultView.ToTable(false, new string[] { a5, b5, c5 });
                ImportToCSV(newTable5, othersPath + "\\" + "属性表.csv");

                for (int i = 0;i<nlayers;i++)
                {
                    Geoprocessor gp = new Geoprocessor();
                    gp.OverwriteOutput = true;
                    gp.SetEnvironmentValue("outputCoordinateSystem", text);
                    ESRI.ArcGIS.DataManagementTools.DeleteField grid12 = new ESRI.ArcGIS.DataManagementTools.DeleteField();
                    grid12.in_table = gridPath + "\\" + "layer" + i + ".shp";
                    grid12.drop_field = "宽度;高度";
                    gp.Execute(grid12, null);

                    ESRI.ArcGIS.DataManagementTools.AddGeometryAttributes grid13 = new ESRI.ArcGIS.DataManagementTools.AddGeometryAttributes();
                    grid13.Input_Features = gridPath + "\\" + "layer" + i + ".shp";
                    grid13.Geometry_Properties = "CENTROID";
                    grid13.Length_Unit = "METERS";
                    grid13.Area_Unit = "SQUARE_METERS";
                    gp.Execute(grid13, null);


                }

                //给Layerpropety添加默认字段
                string connStr = @"Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:DataBase Password=;Data Source=" + mdbPath;
                OleDbConnection tempconn = new OleDbConnection(connStr);
                OleDbCommand cmd = new OleDbCommand("delete * From LayerProperty", tempconn);
                tempconn.Open();
                cmd.ExecuteNonQuery();

                string str1 = "Alter TABLE [LayerProperty] Alter COLUMN ID COUNTER (1, 1)";
                OleDbCommand cmd1 = new OleDbCommand(str1, tempconn);
                cmd1.ExecuteNonQuery();


                for (int i =0;i<nlayers;i++)
                {
                    string layer1 = "INSERT INTO [LayerProperty] ([Layer],[Type],[Horizontal Anisotropy],[Vertical Anisotropy],[Transmissivity],[Leakance],[Storage Coefficient],[Interbed Storage],[Density]) values('" + ("layer" + i) + "', '"+ "Confined" + "', '" + "1" + "', '" + "VK" + "', '" + "Calculated" + "', '" + "Calculated" + "', '" + "User Specified" + "', '" + "0" + "', '" + "0" + "')";
                    OleDbCommand oleDbCommand = new OleDbCommand(layer1, tempconn);
                    oleDbCommand.ExecuteNonQuery();
                }



                //删除行列层数
                string del1 = "delete * From Others where [Coord] = '行号'";
                string del2 = "delete * From Others where [Coord] = '列号'";
                string del3 = "delete * From Others where [Coord] = '层数'";
                string del4 = "delete * From Others where [Coord] = '起始点X'";
                string del5 = "delete * From Others where [Coord] = '起始点Y'";
                string del6 = "delete * From Others where [Coord] = '顶板状态'";
                string del7 = "delete * From Others where [Coord] = 'IBOUND状态'";
                OleDbCommand DEL1 = new OleDbCommand(del1, tempconn);
                OleDbCommand DEL2 = new OleDbCommand(del2, tempconn);
                OleDbCommand DEL3 = new OleDbCommand(del3, tempconn);
                OleDbCommand DEL4 = new OleDbCommand(del4, tempconn);
                OleDbCommand DEL5 = new OleDbCommand(del5, tempconn);
                OleDbCommand DEL6 = new OleDbCommand(del6, tempconn);
                OleDbCommand DEL7 = new OleDbCommand(del7, tempconn);
                DEL1.ExecuteNonQuery();
                DEL2.ExecuteNonQuery();
                DEL3.ExecuteNonQuery();
                DEL4.ExecuteNonQuery();
                DEL5.ExecuteNonQuery();
                DEL6.ExecuteNonQuery();
                DEL7.ExecuteNonQuery();




                //把行列数添加进others
                string sql2 = "INSERT INTO [Others] ([Coord],[Text]) values('" + "行号" + "', '" + nrows + "')";
                OleDbCommand oleDbCommand2 = new OleDbCommand(sql2, tempconn);
                oleDbCommand2.ExecuteNonQuery();

                string sql3 = "INSERT INTO [Others] ([Coord],[Text]) values('" + "列号" + "', '" + ncolums + "')";
                OleDbCommand oleDbCommand3 = new OleDbCommand(sql3, tempconn);
                oleDbCommand3.ExecuteNonQuery();

                string sql4 = "INSERT INTO [Others] ([Coord],[Text]) values('" + "层数" + "', '" + nlayers + "')";
                OleDbCommand oleDbCommand4 = new OleDbCommand(sql4, tempconn);
                oleDbCommand4.ExecuteNonQuery();

                string sql5 = "INSERT INTO [Others] ([Coord],[Text]) values('" + "起始点X" + "', '" + coordx + "')";
                OleDbCommand oleDbCommand5 = new OleDbCommand(sql5, tempconn);
                oleDbCommand5.ExecuteNonQuery();

                string sql6 = "INSERT INTO [Others] ([Coord],[Text]) values('" + "起始点Y" + "', '" + coordy + "')";
                OleDbCommand oleDbCommand6 = new OleDbCommand(sql6, tempconn);
                oleDbCommand6.ExecuteNonQuery();

                string sql7 = "INSERT INTO [Others] ([Coord],[Text]) values('" + "顶板状态" + "', '" + 0 + "')";
                OleDbCommand oleDbCommand7 = new OleDbCommand(sql7, tempconn);
                oleDbCommand7.ExecuteNonQuery();

                string sql8 = "INSERT INTO [Others] ([Coord],[Text]) values('" + "IBOUND状态" + "', '" + 0 + "')";
                OleDbCommand oleDbCommand8 = new OleDbCommand(sql8, tempconn);
                oleDbCommand8.ExecuteNonQuery();

                string del = "Alter TABLE [Others] Alter COLUMN ID COUNTER (1, 1)";               
                OleDbCommand cmd3 = new OleDbCommand(del, tempconn);
                cmd3.ExecuteNonQuery();
                mxd.UpdateContents();


                


                MessageBox.Show(nlayers+"层网格创建成功！", "提示", MessageBoxButtons.OK);

            }
            catch (Exception ex)
            {
                //MessageBox.Show("请输入有效数据！", "提示", MessageBoxButtons.OK);
            }
            this.Close();
        }
        public string GetActiveDocumentPath(IMapDocument mapDocument)
        {
            IDocumentInfo2 docInfo = mapDocument as IDocumentInfo2;
            return docInfo.Path;
        }

        private static System.Data.DataTable GetAttributesTable(IFeatureClass pFeatureClass)
        {
            string geometryType = string.Empty;
            if (pFeatureClass.ShapeType == esriGeometryType.esriGeometryPoint)
            {
                geometryType = "点";
            }
            if (pFeatureClass.ShapeType == esriGeometryType.esriGeometryMultipoint)
            {
                geometryType = "点集";
            }
            if (pFeatureClass.ShapeType == esriGeometryType.esriGeometryPolyline)
            {
                geometryType = "折线";
            }
            if (pFeatureClass.ShapeType == esriGeometryType.esriGeometryPolygon)
            {
                geometryType = "面";
            }

            // 字段集合
            IFields pFields = pFeatureClass.Fields;
            int fieldsCount = pFields.FieldCount;

            // 写入字段名
            System.Data.DataTable dataTable = new System.Data.DataTable();
            for (int i = 0; i < fieldsCount; i++)
            {
                dataTable.Columns.Add(pFields.get_Field(i).Name);
            }

            // 要素游标
            IFeatureCursor pFeatureCursor = pFeatureClass.Search(null, true);
            IFeature pFeature = pFeatureCursor.NextFeature();
            if (pFeature == null)
            {
                return dataTable;
            }

            // 获取MZ值
            IMAware pMAware = pFeature.Shape as IMAware;
            IZAware pZAware = pFeature.Shape as IZAware;
            if (pMAware.MAware)
            {
                geometryType += " M";
            }
            if (pZAware.ZAware)
            {
                geometryType += "Z";
            }

            // 写入字段值
            while (pFeature != null)
            {
                DataRow dataRow = dataTable.NewRow();
                for (int i = 0; i < fieldsCount; i++)
                {
                    if (pFields.get_Field(i).Type == esriFieldType.esriFieldTypeGeometry)
                    {
                        dataRow[i] = geometryType;
                    }
                    else
                    {
                        dataRow[i] = pFeature.get_Value(i).ToString();
                    }
                }
                dataTable.Rows.Add(dataRow);
                pFeature = pFeatureCursor.NextFeature();
            }

            // 释放游标
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pFeatureCursor);
            return dataTable;
        }

        public static DataTable DeleteSameRow(DataTable dt, string Field)
        {
            ArrayList indexList = new ArrayList();
            // 找出待删除的行索引   
            for (int i = 0; i < dt.Rows.Count - 1; i++)
            {
                if (!IsContain(indexList, i))
                {
                    for (int j = i + 1; j < dt.Rows.Count; j++)
                    {
                        if (dt.Rows[i][Field].ToString() == dt.Rows[j][Field].ToString())
                        {
                            indexList.Add(j);
                        }
                    }
                }
            }
            indexList.Sort();
            // 排序
            for (int i = indexList.Count - 1; i >= 0; i--)// 根据待删除索引列表删除行  
            {
                int index = Convert.ToInt32(indexList[i]);
                dt.Rows.RemoveAt(index);
            }
            return dt;
        }

        public static bool IsContain(ArrayList indexList, int index)
        {
            for (int i = 0; i < indexList.Count; i++)
            {
                int tempIndex = Convert.ToInt32(indexList[i]);
                if (tempIndex == index)
                {
                    return true;
                }
            }
            return false;
        }


        public static void ImportToCSV(System.Data.DataTable dt, string fileName)
        {
            FileStream fs = null;
            StreamWriter sw = null;
            try
            {
                fs = new FileStream(fileName, FileMode.Create, FileAccess.Write);
                sw = new StreamWriter(fs, Encoding.Default);
                string head = "";
                //拼接列头
                for (int cNum = 0; cNum < dt.Columns.Count; cNum++)
                {
                    head += dt.Columns[cNum].ColumnName + ",";
                }
                //csv文件写入列头
                sw.WriteLine(head);
                string data = "";
                //csv写入数据
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string data2 = string.Empty;
                    //拼接行数据
                    for (int cNum1 = 0; cNum1 < dt.Columns.Count; cNum1++)
                    {
                        data2 = data2 + dt.Rows[i][dt.Columns[cNum1].ColumnName] + ",";
                    }
                    bool flag = data != data2;
                    if (flag)
                    {
                        sw.WriteLine(data2);
                    }
                    data = data2;

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("导出错误！");
                return;
            }
            finally
            {
                if (sw != null)
                {
                    sw.Close();
                }
                if (fs != null)
                {
                    fs.Close();
                }
                sw = null;
                fs = null;
            }
        }
    }
}
