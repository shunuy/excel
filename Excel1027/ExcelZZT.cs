using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Windows;


namespace Excel1027
{
    class ExcelZZT
    {
        Microsoft.Office.Interop.Excel.Application ThisApplication = null;
        Microsoft.Office.Interop.Excel.Workbooks m_objBooks = null;
        Microsoft.Office.Interop.Excel._Workbook ThisWorkbook = null;

        Microsoft.Office.Interop.Excel.Worksheet xlSheet = null;

        /**/
        /// <summary>
        /// 用生成的随机数作数据
        /// </summary>
        private void LoadData()
        {
            Random ran = new Random();
            //在Excel的工作表的单元中加入文本内容
            //xlSheet、Cels【r，C】= Vaule；
            //其中：r为行数，C为列数。Vaule为单元值。
            xlSheet.Cells[1, 2] = "工业用电";
            xlSheet.Cells[1, 3] = "商业用电";
            for (int i = 1; i <= 12; i++)
            {
                xlSheet.Cells[i+1, 1] = i.ToString() + "月";
                xlSheet.Cells[i+1, 2] = ran.Next(2000).ToString();
                xlSheet.Cells[i+1, 3] = ran.Next(1500).ToString();
            }

        }
        /**/
        /// <summary>
        /// 删除多余的Sheet
        /// </summary>
        private void DeleteSheet()
        {
            foreach (Microsoft.Office.Interop.Excel.Worksheet ws in ThisWorkbook.Worksheets)
                if (ws != ThisApplication.ActiveSheet)
                {
                    ws.Delete();
                }
            foreach (Microsoft.Office.Interop.Excel.Chart cht in ThisWorkbook.Charts)
                cht.Delete();

        }
        /**/
        /// <summary>
        /// 创建一个Sheet，用来存数据
        /// </summary>
        private void AddDatasheet()
        {
            //新建一个工作簿
            xlSheet = (Microsoft.Office.Interop.Excel.Worksheet)ThisWorkbook.
                Worksheets.Add(Type.Missing, ThisWorkbook.ActiveSheet,
                Type.Missing, Type.Missing);

            xlSheet.Name = "数据";
        }
        /**/
        /// <summary>
        /// 创建统计图         
        /// </summary>
        private void CreateChart()
        {
            //图表对象
            Microsoft.Office.Interop.Excel.Chart xlChart = (Microsoft.Office.Interop.Excel.Chart)ThisWorkbook.Charts.
                Add(Type.Missing, xlSheet, Type.Missing, Type.Missing);

            //定义一个工作域(初始化为第一行第一列)
            Microsoft.Office.Interop.Excel.Range cellRange = (Microsoft.Office.Interop.Excel.Range)xlSheet.Cells[1, 1];

            //使用图表向导生成图表
            xlChart.ChartWizard(cellRange.CurrentRegion,//选定区域
                Microsoft.Office.Interop.Excel.XlChartType.xl3DLine,//图表类型：立体直方图[柱形图( xl3DColumn、xlColumnC1ustered),圆饼图(xl3DPie),折线图(xl3DLine),条形图(xlBarClustered)]
                Type.Missing,//内置自动套用格式的选项编号。可以是一个 1 到 10 之间的数（取决于库类型）。如果省略此参数，则 Excel 根据库类型和数据源选择默认值。
                Microsoft.Office.Interop.Excel.XlRowCol.xlColumns, //指定每个系列的数据是按行绘制还是按列绘制。就是用数据表横行的做横坐标还是用数据表的列做横坐标
                1,//一个整数，指定源范围中包含类别标签的行数或列数。 
                1,//一个整数，指定源范围中包含系列标签的行数或列数。数据列的列头作为系列名
                true,//为 true 时包含图例
                "深圳2010年工业用电量与农业用电量比较", //Chart 控件标题文本。 
                "月份", //分类轴标题文本
                "用电量",//数值轴标题文本
                ""//三维图表的系列轴标题或二维图表的第二个数值轴标题
                );

            xlChart.Name = "统计";

            //获取单个图表组（ChartGroup 对象）或图表中所有图表组的集合（ChartGroups 对象)
            Microsoft.Office.Interop.Excel.ChartGroup grp = (Microsoft.Office.Interop.Excel.ChartGroup)xlChart.ChartGroups(1);
            grp.GapWidth = 20;
            //grp.VaryByCategories = true;

            //让Chart的条目的显示形状变成圆柱形，并给它们显示加上数据标签
            Microsoft.Office.Interop.Excel.Series s = (Microsoft.Office.Interop.Excel.Series)grp.SeriesCollection(1);
            s.BarShape = XlBarShape.xlCylinder;
            s.HasDataLabels = true;

            Microsoft.Office.Interop.Excel.Series s1 = (Microsoft.Office.Interop.Excel.Series)grp.SeriesCollection(2);
            s1.BarShape = XlBarShape.xlBox;
            s1.HasDataLabels = true;

            //设置统计图的标题和图例的显示
            xlChart.Legend.Position = XlLegendPosition.xlLegendPositionTop;
            xlChart.ChartTitle.Font.Size = 24;
            xlChart.ChartTitle.Shadow = true;
            xlChart.ChartTitle.Border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            //设置两个轴的属性，Excel.XlAxisType.xlValue对应的是Y轴，Excel.XlAxisType.xlCategory对应的是X轴
            Microsoft.Office.Interop.Excel.Axis valueAxis = (Microsoft.Office.Interop.Excel.Axis)xlChart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue, XlAxisGroup.xlPrimary);
            valueAxis.AxisTitle.Orientation = -90;

            Microsoft.Office.Interop.Excel.Axis categoryAxis = (Microsoft.Office.Interop.Excel.Axis)xlChart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
            categoryAxis.AxisTitle.Font.Name = "MS UI Gothic";
        }


        public void CreatZZT()
        {
            try
            {
                //定义一个Excel应用程序
                ThisApplication = new Microsoft.Office.Interop.Excel.Application();
                //定义一个工作簿
                m_objBooks = (Microsoft.Office.Interop.Excel.Workbooks)ThisApplication.Workbooks;
                ThisWorkbook = (Microsoft.Office.Interop.Excel._Workbook)(m_objBooks.Add(Type.Missing));

                ThisApplication.DisplayAlerts = false;

                this.DeleteSheet();
                this.AddDatasheet();
                this.LoadData();

                CreateChart();

                ThisWorkbook.SaveAs("E:\\Book2.xls", Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ThisWorkbook.Close(Type.Missing, Type.Missing, Type.Missing);
                ThisApplication.Workbooks.Close();

                ThisApplication.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ThisWorkbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ThisApplication);
                ThisWorkbook = null;
                ThisApplication = null;
            }
        }
    }
}
