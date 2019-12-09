using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Web;
using System.Web.SessionState;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using System.Text;

namespace DataMining.Common
{
    /// <summary>
    /// ExportFiles 的摘要说明。
    /// 作用：把DataSet数据集内数据转化为Excel、Word文件
    /// 描述：这些关于Excel、Word的导出方法，基本可以实现日常须要，其中有些方法可以把数据导出后
    ///       生成Xml格式，再导入数据库！有些屏蔽内容没有去掉，保留下来方便学习参考用之。   
    /// 备注：请引用Office相应COM组件，导出Excel对象的一个方法要调用其中的一些方法和属性。
    /// </summary>

    public class ExportFiles  
    {

        /// <summary>
        /// 
        /// </summary>
        /// 
        #region  //构造函数

        public ExportFiles()
        {
            //
            // TODO: 在此处添加构造函数逻辑
            //
        }
        #endregion

        public void writetext(string content,string pathfile){
            System.IO.File.WriteAllText(pathfile, content, System.Text.Encoding.GetEncoding("utf-8"));
        }

        /// <summary>
        /// 调用Excel.dll导出Excel文件
        /// </summary>
        /// <param name="ds"></param>
        /// 
        /*
        #region  
        // 调用Excel.dll导出Excel文件
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ds">DataSet数据庥</param>
        /// <param name="Duser">登录用户（如session["username"].Tostring()）可为null或空</param>
        /// <param name="titlename">添加一个报表标题</param>
        /// <param name="filepath">指定文件在服务器上的存放地址（如：Server.MapPath(".")）可为null或空</param>
        /// 
        /// 为什么在这里设置个filepath?
        /// 原因如下：filepath接收的内容是Server.MapPath(".")这个参数值,这个值在这本类中对
        /// System.Web.HttpServerUtility HServer=new System.Web.HttpServerUtility()的引用出错，因为没有继承Page类
        /// 所以只能以传值的形式调用了，你有好的办法可以修改，那就再好不过了！
        public void DataSetToExcel(DataSet ds, string Duser, string titlename, string filepath)
        {
            //Microsoft.Office.Interop.Owc11() appowc=new Microsoft.Office.Interop.Owc11();
            OWC.Owc11.SpreadsheetClass xlsheet = new Microsoft.Office.Interop.Owc11.SpreadsheetClass();

            #region //屏蔽内容
            ///本来想用下面的这个办法实现的，可在IIS中必须设置相关的权限
            ///所以就放弃了，把代码设置为屏蔽内容，供学习参考！
            ///
            //   Microsoft.Office.Interop.Excel.Application app=new Microsoft.Office.Interop.Excel.Application();
            //   
            //   if(app==null)
            //   {
            //               throw new Exception("系统调用错误（Excel.dll）");
            //   }
            //   app.Application.Workbooks.Add(true);
            //   WorkbookClass oBook=new WorkbookClass();
            //            WorksheetClass oSheet=new WorksheetClass();
            //   
            //   //定义表对象与行对象，同时用DataSet对其值进行初始化
            //   System.Data.DataTable dt=ds.Tables[0];
            //      oSheet.get_Range(app.Cells[1,1],app.Cells[10,15]).HorizontalAlignment=OWC.Owc11.XlHAlign.xlHAlignCenter;
            //   DataRow[] myRow=dt.Select();
            //   int i=0;
            //   int cl=dt.Columns.Count;
            //   //取得数据表各列标题
            //   for(i=0;i<cl;i++)
            //   {   
            //    app.Cells[1,i+1]=dt.Columns[i].Caption.ToString();
            //    //app.Cells.AddComment(dt.Columns[i].Caption.ToString());
            //    //oSheet.Cells.AddComment(dt.Columns[i].Caption.ToString());
            //    //app.Cells=dt.Columns[i].Caption.ToString();
            //    //oSheet.get_Range(app.Cells,app.Cells).HorizontalAlignment=
            //    //app.Cells.AddComment=dt.Columns[i].ToString();
            //    //oSheet.get_Range(app.Cells,app.Cells).HorizontalAlignment=
            //   }
            #endregion

            //定义表对象与行对象，同时用DataSet对其值进行初始化
            System.Data.DataTable dt = ds.Tables[0];
            DataRow[] myRow = dt.Select();
            int i = 0;
            int col = 1;
            int colday = col + 1;
            int colsecond = colday + 1;
            int colnumber = colsecond + 1;
            int cl = dt.Columns.Count;
            string userfile = null;
            //合并单元格
            xlsheet.get_Range(xlsheet.Cells[col, col], xlsheet.Cells[col, cl]).set_MergeCells(true);

            //添加标题名称
            if (titlename == "" || titlename == null)
                xlsheet.ActiveSheet.Cells[col, col] = "添加标题处（高级报表）";
            else
                xlsheet.ActiveSheet.Cells[col, col] = titlename.Trim();

            //判断传值user是否为空
            if (Duser == "" || Duser == null)
                userfile = "DFSOFT";
            else
                userfile = Duser;

            //设置标题大小
            xlsheet.get_Range(xlsheet.Cells[col, col], xlsheet.Cells[col, cl]).Font.set_Size(13);

            //加粗标题
            xlsheet.get_Range(xlsheet.Cells[col, col], xlsheet.Cells[col, cl]).Font.set_Bold(true);
            xlsheet.get_Range(xlsheet.Cells[colsecond, col], xlsheet.Cells[colsecond, cl]).Font.set_Bold(true);
            //设置标题水平居中
            xlsheet.get_Range(xlsheet.Cells, xlsheet.Cells).set_HorizontalAlignment(OWC.Owc11.XlHAlign.xlHAlignCenter);


            //设置单元格宽度
            //xlsheet.get_Range(xlsheet.Cells,xlsheet.Cells).set_ColumnWidth(9);

            xlsheet.get_Range(xlsheet.Cells[colday, col], xlsheet.Cells[colday, cl]).set_MergeCells(true);
            xlsheet.ActiveSheet.Cells[colday, col] = "日期：" + DateTime.Now.Year.ToString() + "年" + DateTime.Now.Month.ToString() + "月" + DateTime.Now.Day.ToString() + "日  ";
            xlsheet.get_Range(xlsheet.Cells[colday, col], xlsheet.Cells[colday, cl]).set_HorizontalAlignment(OWC.Owc11.XlHAlign.xlHAlignRight);


            //取得数据表各列标题，各标题之间以\t分割，最后一个列标题后加回车符
            for (i = 0; i < cl; i++)
            {
                xlsheet.ActiveSheet.Cells[colsecond, i + 1] = dt.Columns[i].Caption.ToString();
            }
            //逐行处理数据
            foreach (DataRow row in myRow)
            {
                //当前数据写入
                for (i = 0; i < cl; i++)
                {
                    xlsheet.ActiveSheet.Cells[colnumber, i + 1] = row[i].ToString().Trim();
                }
                colnumber++;
            }
            //设置边框线
            xlsheet.get_Range(xlsheet.Cells[colsecond, col], xlsheet.Cells[colnumber - 1, cl]).Borders.set_LineStyle(OWC.Owc11.XlLineStyle.xlContinuous);

            try
            {
                //xlsheet.get_Range(xlsheet.Cells[2,1],xlsheet.Cells[8,15]).set_NumberFormat("￥#,##0.00");
                // System.Web.HttpServerUtility HServer=new System.Web.HttpServerUtility();

                //HServer.MapPath(".")+"//testowc.xls";
                xlsheet.Export(filepath + "//exportfiles//~$" + userfile + ".xls", OWC.Owc11.SheetExportActionEnum.ssExportActionNone, OWC.Owc11.SheetExportFormat.ssExportXMLSpreadsheet);


            }
            catch (Exception e)
            {
                throw new Exception("系统调用错误或有打开的Excel文件！" + e);
            }

            //Web页面定义
            HttpResponse resp;
            resp = HttpContext.Current.Response;
            resp.ContentEncoding = System.Text.Encoding.GetEncoding("GB2312");
            resp.AppendHeader("Content-disposition", "attachment;filename=" + userfile + ".xls");
            resp.ContentType = "application/ms-excel";

            string path = filepath + "//exportfiles//~$" + userfile + ".xls";
            System.IO.FileInfo file = new FileInfo(path);
            resp.Clear();
            resp.AddHeader("content-length", file.Length.ToString());
            resp.WriteFile(file.FullName);
            resp.End();
        }
        #endregion
        */
        /// <summary>
        /// 导出Excel文件类
        /// </summary>
        /// <param name="ds"></param>
        /// <param name="FileName"></param>
        ///
        #region  
        
        //导出Excel文件类
        public void DataSetToExcel(DataSet ds, string FileName)
        {
            try
            {
                //Web页面定义
                //System.Web.UI.Page mypage=new System.Web.UI.Page();
                HttpResponse resp;
                resp = HttpContext.Current.Response;
                resp.ContentEncoding = System.Text.Encoding.GetEncoding("GB2312");
                //resp.AppendHeader("Content-disposition", "attachment;filename=" + FileName + ".xls");
                resp.ContentType = "application/ms-excel";
                //变量定义
                string colHeaders = null;
                string Is_item = null;

                //显示格式定义////////////////


                //文件流操作定义
                //  FileStream fs=new FileStream(FileName,FileMode.Create,FileAccess.Write);
                //StreamWriter sw=new StreamWriter(fs,System.Text.Encoding.GetEncoding("GB2312"));

                StringWriter sfw = new StringWriter();
                //定义表对象与行对象，同时用DataSet对其值进行初始化
                System.Data.DataTable dt = ds.Tables[0];
                DataRow[] myRow = dt.Select();
                int i = 0;
                int cl = dt.Columns.Count;

                //取得数据表各列标题，各标题之间以\t分割，最后一个列标题后加回车符
                for (i = 0; i < cl; i++)
                {
                    //if(i==(cl-1))  //最后一列，加\n
                    // colHeaders+=dt.Columns[i].Caption.ToString();
                    //else
                    colHeaders += dt.Columns[i].Caption.ToString() + "\t";
                }
                sfw.WriteLine(colHeaders);
                //sw.WriteLine(colHeaders);

                //逐行处理数据
                foreach (DataRow row in myRow)
                {
                    //当前数据写入
                    for (i = 0; i < cl; i++)
                    {
                        //if(i==(cl-1))
                        //   Is_item+=row[i].ToString()+"\n";
                        //else
                        Is_item += row[i].ToString() + "\t";
                    }
                    sfw.WriteLine(Is_item);
                    //sw.WriteLine(Is_item);
                    Is_item = null;
                }
                resp.Write(sfw);
                //resp.Clear();
                //resp.End();

              
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        #endregion


        /// <summary>
        /// 数据集转换，即把DataSet转换为Excel对象
        /// </summary>
        /// <param name="ds"></param>
        /// <param name="FileName"></param>
        /// <param name="titlename"></param>
        /// 
        #region    
        //运用html+css生成Excel

        public void DataSetToExcel(DataSet ds, string FileName, string titlename)
        {
            string ExportFileName = null;
            if (FileName == null || FileName == "")
                ExportFileName = "Excel";
            else
                ExportFileName = FileName;

            if (titlename == "" || titlename == null)
                titlename = "添加标题处（高级报表）";


            //定义表对象与行对象，同时用DataSet对其值进行初始化
            System.Data.DataTable dt = ds.Tables[0];
            DataRow[] myRow = dt.Select();
            int i = 0;
            int cl = dt.Columns.Count;




            //Web页面定义 
            StringWriter strHTML = new StringWriter();
            HttpResponse resp = new HttpResponse(strHTML);
            resp.Clear();
            resp.ContentEncoding = System.Text.Encoding.GetEncoding("utf-8");
            resp.AppendHeader("Content-disposition", "attachment;filename=" + ExportFileName + ".xls");
            resp.ContentType = "application/vnd.ms-excel";

            string BeginTab = "<table border='0' cellpadding='0' cellspacing='0' style='border-right:#000000 0.1pt solid;border-top:#000000 0.1pt solid;'>";
            string EndTab = "</table>";
            StringBuilder FileIO = new StringBuilder();
            StringBuilder MainIO = new StringBuilder();
            string TitleTab = "<tr><td colspan='" + cl + "' style='font-size:30px;' align='center'><b>" + titlename + "</b></td></tr><tr><td colspan='" + cl + "' align='right' style='font-size:15px;'>" + DateTime.Now.Year.ToString() + "年" + DateTime.Now.Month.ToString() + "月" + DateTime.Now.Day.ToString() + "日&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>";
            string BeginTr = "<tr>";
            string EndTr = "</tr>";
            for (i = 0; i < cl; i++)
            {
                FileIO.Append( "<td style='border-left:#000000 0.1pt solid; border-bottom:#000000 1.0pt solid; font-size:15px;' align='center'><b>" + dt.Columns[i].Caption.ToString() + "</b></td>");

            }
            FileIO.Append( BeginTr.ToString() + FileIO.ToString() + EndTr.ToString());

            //逐行处理数据
            foreach (DataRow row in myRow)
            {
                StringBuilder OutIO = new StringBuilder();
                //当前数据写入
                for (i = 0; i < cl; i++)
                {
                    OutIO.Append( "<td style='border-left:#000000 0.1pt solid; border-bottom:#000000 1.0pt solid; font-size:15px;' align='center'>" + row[i].ToString() + "</td>");

                }
                MainIO.Append(BeginTr.ToString() + OutIO.ToString() + EndTr.ToString());
            }

            FileIO.Append( "<center><table>" + TitleTab.ToString() + "<tr>" + BeginTab.ToString() + FileIO.ToString() + MainIO.ToString() + EndTab.ToString() + "</tr></table></center>");

            resp.Write(FileIO.ToString());
            writetext(strHTML.ToString(), FileName);



        }
        #endregion

        /// <summary>
        /// 导出Word文件类
        /// </summary>
        /// <param name="ds"></param>
        /// <param name="FileName"></param>
        ///
        #region 
        //导出Word文件类

        public void DataSetToWord(DataSet ds, string FileName)
        {
            //try
            //{

            //StringWriter strHTML = new StringWriter();
            //System.Web.UI.Page myPage = new Page();   //System.Web.UI.Page中有个Server对象，我们要利用一下它
            //myPage.Server.Execute(strUrl, strHTML);    //将asp_net.aspx将在客户段显示的html内容读到了strHTML中
            //StreamWriter sw = new StreamWriter(strSavePath + strSaveFile, true, System.Text.Encoding.GetEncoding("GB2312"));
            ////新建一个文件Test.htm，文件格式为GB2312
            //sw.Write(strHTML.ToString());             //将strHTML中的字符写到Test.htm中
            //strHTML.Close();                          //关闭StringWriter 
            //sw.Close();                                    //关闭StreamWriter 
            //return true;


                //Web页面定义
                //System.Web.UI.Page mypage = new System.Web.UI.Page();
                 //HttpResponse resp;
                 //resp = HttpContext.Current.Response;
                StringWriter strHTML = new StringWriter();
                HttpResponse resp= new HttpResponse(strHTML);
 
                resp.Clear();
                resp.Buffer = true;
                resp.Charset = "utf-8";
                resp.ContentEncoding = System.Text.Encoding.GetEncoding("utf-8");
               // resp.AppendHeader("Content-disposition", "attachment;filename=" + FileName + ".doc");
                resp.ContentType = "application/ms-word";
                //变量定义
                string colHeaders = null;
                string Is_item = null;

                //显示格式定义////////////////


                //文件流操作定义
                //  FileStream fs=new FileStream(FileName,FileMode.Create,FileAccess.Write);
                //StreamWriter sw=new StreamWriter(fs,System.Text.Encoding.GetEncoding("GB2312"));

                StringWriter sfw = new StringWriter();
                //定义表对象与行对象，同时用DataSet对其值进行初始化
                System.Data.DataTable dt = ds.Tables[0];
                DataRow[] myRow = dt.Select();
                int i = 0;
                int cl = dt.Columns.Count;

                //取得数据表各列标题，各标题之间以\t分割，最后一个列标题后加回车符
                for (i = 0; i < cl; i++)
                {
                    //if(i==(cl-1))  //最后一列，加\n
                    // colHeaders+=dt.Columns[i].Caption.ToString();
                    //else
                    colHeaders += dt.Columns[i].Caption.ToString() + "\t";
                }
                sfw.WriteLine(colHeaders);
                //sw.WriteLine(colHeaders);

                //逐行处理数据
                foreach (DataRow row in myRow)
                {
                    //当前数据写入
                    for (i = 0; i < cl; i++)
                    {
                        //if(i==(cl-1))
                        //   Is_item+=row[i].ToString()+"\n";
                        //else
                        Is_item += row[i].ToString() + "\t";
                    }
                    sfw.WriteLine(Is_item);
                    //sw.WriteLine(Is_item);
                    Is_item = null;
                }
                resp.Write(sfw);
                //resp.Clear();
                //resp.End();
                writetext(strHTML.ToString(), FileName);

            //}
            //catch (Exception e)
            //{
            //    throw e;
            //}

        }
        #endregion

        /// <summary>
        /// 数据集转换，即把DataSet转换为Word对象
        /// </summary>
        /// <param name="ds"></param>
        /// <param name="titlename"></param>
        /// 
        #region  
        // 运行html+css生成Word文件

        public void DataSetToWord(DataSet ds, string FileName, string titlename)
        {
            //调用Office
            //备注：速度太慢放弃应用此方法
            //OWC.Word.Application oWord=new OWC.Word.ApplicationClass();
            //OWC.Word._Document oDoc=new OWC.Word.DocumentClass();

            string ExportFileName = null;
            if (FileName == null || FileName == "")
                ExportFileName = "DFSOFT";
            else
                ExportFileName = FileName;

            if (titlename == "" || titlename == null)
                titlename = "添加标题处（高级报表）";


            //定义表对象与行对象，同时用DataSet对其值进行初始化
            System.Data.DataTable dt = ds.Tables[0];
            DataRow[] myRow = dt.Select();
            int i = 0;
            int cl = dt.Columns.Count;



            #region
            //   string FileTitle="<center><table><tr><td><b>报表测试</b></td></tr></table>"+"\n";
            //   string EndFile="</center>";
            //   //Web页面定义 
            StringWriter strHTML = new StringWriter();
            HttpResponse resp = new HttpResponse(strHTML);
            resp.Clear();
            resp.ContentEncoding = System.Text.Encoding.GetEncoding("utf-8");
            resp.AppendHeader("Content-disposition", "attachment;filename=" + ExportFileName + ".doc");
            resp.ContentType = "application/vnd.ms-word";
            //   System.IO.StringWriter oSW=new StringWriter();
            //      System.Web.UI.HtmlTextWriter oHW=new System.Web.UI.HtmlTextWriter(oSW);
            //   System.Web.UI.WebControls.DataGrid oDG=new System.Web.UI.WebControls.DataGrid();
            //            oDG.DataSource=ds.Tables[0];
            //   oDG.DataBind();
            //   oDG.RenderControl(oHW);
            //   resp.Write(FileTitle.ToString()+oSW.ToString()+EndFile.ToString());
            //   resp.End();
            #endregion
            string BeginTab = "<table border='0' cellpadding='0' cellspacing='0' style='border-right:#000000 0.1pt solid;border-top:#000000 0.1pt solid;'>";
            string EndTab = "</table>";
            StringBuilder FileIO = new StringBuilder();
            StringBuilder MainIO = new StringBuilder();
            string TitleTab = "<tr><td style='font-size:13px;' align='center'><b>" + titlename + "</b></td></tr><tr><td align='right' style='font-size:15px;'>" + DateTime.Now.Year.ToString() + "年" + DateTime.Now.Month.ToString() + "月" + DateTime.Now.Day.ToString() + "日&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>";
            string BeginTr = "<tr>";
            string EndTr = "</tr>";
            for (i = 0; i < cl; i++)
            {
                FileIO.Append("<td style='border-left:#000000 0.1pt solid; border-bottom:#000000 1.0pt solid; font-size:15px;' align='center'><b>" + dt.Columns[i].Caption.ToString() + "</b></td>");

            }
            FileIO.Append(BeginTr.ToString() + FileIO.ToString() + EndTr.ToString());

            //逐行处理数据
            foreach (DataRow row in myRow)
            {
                StringBuilder OutIO = new StringBuilder();
                //当前数据写入
                for (i = 0; i < cl; i++)
                {
                    OutIO.Append("<td style='border-left:#000000 0.1pt solid; border-bottom:#000000 1.0pt solid; font-size:15px;' align='center'>" + row[i].ToString() + "</td>");

                }
                MainIO.Append(BeginTr.ToString() + OutIO.ToString() + EndTr.ToString());
            }

            FileIO.Append("<center><table>" + TitleTab.ToString() + "<tr>" + BeginTab.ToString() + FileIO.ToString() + MainIO.ToString() + EndTab.ToString() + "</tr></table></center>");

            resp.Write(FileIO.ToString());
            writetext(strHTML.ToString(), FileName);

        }
        #endregion
    }
}


