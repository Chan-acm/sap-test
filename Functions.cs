using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Data.Sql;
using System.Data.SqlTypes;
using System.Data;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Runtime.InteropServices;



namespace ACM_SAPTEST1
{
   public class Functions
    {

        SqlConnection Conn = new SqlConnection(ConfigurationManager.ConnectionStrings["SAPDBConnectionString"].ConnectionString);
        SqlDataAdapter SQLAdapter = new SqlDataAdapter();
        SqlCommand Cmd = new SqlCommand();

        DataTable mDT = new DataTable();
        DataTable mDMSdt = new DataTable();
        
        #region " Export Excel "
        public bool RetriveData(string criservice, [Optional] string crifromdate, [Optional] string critodate, [Optional] string criWhcode, [Optional] string criGrpcode, [Optional] string criItemType, [Optional] string criItmService)
        {
            var query = "";
            string strtblname = "";
            string TDate = critodate.Replace("/", "");
            if (criservice == "SALES_ORDER")
            {
                query = string.Format(@"SELECT DocNum,CASE WHEN DocType = 'I' THEN ' dDocument_Items' ELSE 'dDocument_Service' END AS DocType,
                                       convert(varchar, DocDate, 112) as DocDate, convert(varchar, DocDueDate, 112) as DocDueDate,CardCode,NumAtCard,DocTotal,Ref1,Ref2,Series,convert(varchar, TaxDate, 112) as TaxDate,
										isnull(U_DocNum,'') as U_DocNum,isnull(U_Route,'') As U_Route,isnull(convert(varchar, U_ExpiredDate, 112) ,'') as U_ExpiredDate,
                                        isnull(U_Township,'') as U_Township,isnull(U_TotalQuantity,0.00) as U_TotalQuantity,isnull(U_TotalMSU,'') as U_TotalMSU,
                                        isnull(U_RequestType,'') as U_RequestType,isnull(U_UserCode,'') as U_UserCode,isnull(U_OM,'') as U_OM,
                                        isnull(U_TR_Status,'') as U_TR_Status,isnull(U_TotalWeight,'') as U_TotalWeight,isnull(U_VendorCode,'') as U_VendorCode,
                                        isnull(U_PRTotalFC,'') as U_PRTotalFC,isnull(U_VANCode,'') as U_VANCode,isnull(U_PickStatus,'') as U_PickStatus,
                                        isnull(U_BaseEntry,'') As U_BaseEntry,isnull(U_BaseType,'') as U_BaseType,isnull(U_Consignment,'') as U_Consignment,
                                        isnull(U_PromotionCode,'') As U_PromotionCode,isnull(U_DiscPrcnt,0.00) as U_DiscPrcnt,isnull(U_BranchEntry,'') as U_BranchEntry,
                                        isnull(U_BranchCode,'') as U_BranchCode,isnull(U_BranchLineNum,0) as U_BranchLineNum,isnull(U_DONumber,'') as U_DONumber,
                                        isnull(U_PDA_OrderID,'') as U_PDA_OrderID
                                        FROM ORDR
                                        WHERE (DocDate >= '" + crifromdate.Replace ("/","") + "' and  DocDate <= '" + critodate.Replace ("/","") + "') AND DocStatus = 'O'");
                // AND DocType = '" + criItemType + "'


                strtblname = criservice + "_Header_" + criItmService + "_" + TDate;
                ExportDTtoExcel(query, strtblname);

                query = string.Format(@"SELECT R.DocNum,R1.LineNum,R1.ItemCode,R1.Dscription,R1.Quantity,convert(varchar, R1.ShipDate, 112) as ShipDate,R1.Price,R1.WhsCode,R1.AcctCode,R1.UseBaseUn,
                                        R1.BaseType,R1.BaseEntry,R1.BaseLine,R1.TaxCode,R1.TaxType,R1.NumPerMsr,R1.LineTotal,
										isnull(R1.U_RecommendPrice,0.00) as U_RecommendPrice,isnull(R1.U_NotApplyPromotion,'') as U_NotApplyPromotion,
                                        isnull(R1.U_PromotionCode,'') as U_PromotionCode,isnull(R1.U_SUFactor,'') as U_SUFactor,isnull(R1.U_MSUOrder,'') as U_MSUOrder,
                                        isnull(R1.U_PromotionLine,'') as U_PromotionLine,isnull(R1.U_OrgPrice,0.00) as U_OrgPrice,isnull(R1.U_DiscPrcnt,0.00) as U_DiscPrcnt,
                                        isnull(R1.U_WHName,'') as U_WHName,isnull(R1.U_BranchLineNum,'') as U_BranchLineNum
                                        FROM RDR1 R1
                                        INNER JOIN ORDR R ON R1.DocEntry = R.DocEntry
                                        WHERE (R.DocDate >= '" + crifromdate.Replace("/", "") + "' and  R.DocDate <= '" + critodate.Replace("/", "") + "') AND R.DocStatus = 'O' ");
                //AND R.DocType = '" + criItmService + "'

                strtblname = criservice + "_Detail_" + criItmService + "_" + TDate;
                ExportDTtoExcel(query, strtblname);


            }
            else
            {
                //do something
            }
            MessageBox.Show(criservice + " Successfully Exported!", "Export DTW", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return true;
        }

        public bool ExportDTtoExcel(string criquery, string tblname)
        {

            Janus.Windows.GridEX.GridEX gr = new Janus.Windows.GridEX.GridEX();

            DataTable dt = SAP_Local_RunQuery(criquery);
            //dt.DefaultView.Sort = "id asc"; // order by id
            dt = dt.DefaultView.ToTable();
            if (dt.Rows.Count > 0)
            {
                gr.DataSource = dt;
                gr.RetrieveStructure();

                foreach (Janus.Windows.GridEX.GridEXColumn col in gr.RootTable.Columns)
                {
                    col.Caption = col.DataMember;
                }

                ExportGridEx(gr, tblname);

            }
            else
            {

                MessageBox.Show("There is no record!", "PHM Export DTW", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            return true;

        }

        public static bool ExportGridEx(Janus.Windows.GridEX.GridEX Grid, string fname)
        {
            Janus.Windows.GridEX.Export.GridEXExporter Export = new Janus.Windows.GridEX.Export.GridEXExporter();
            Export.GridEX = Grid;


            string strpath = @"C:\DTW_Export"; //Application.ExecutablePath + @"\DTW_Export";
            if (Directory.Exists(strpath) == false)
            {
                Directory.CreateDirectory(strpath);
            }


            SaveFileDialog SaveFileDialog1 = new SaveFileDialog();
            SaveFileDialog1.Filter = "Excel files (*.xls,*.xlsx,*.xml)|*.*";
            //SaveFileDialog1.InitialDirectory = @"C:\Export_DTW_Excel";
            SaveFileDialog1.InitialDirectory = strpath;
            SaveFileDialog1.FileName = fname + ".xls";
            SaveFileDialog1.DefaultExt = "xml";
            if (SaveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string ReportFile = SaveFileDialog1.FileName;
                FileStream stream = new FileStream(ReportFile, FileMode.Create);

                try
                {
                    Export.IncludeHeaders = true;
                    Export.ExportMode = Janus.Windows.GridEX.ExportMode.AllRows;
                    Export.IncludeExcelProcessingInstruction = true;
                    Export.IncludeFormatStyle = true;
                    Export.Export(stream);
                    stream.Flush();
                    int rowcount = Export.GridEX.RowCount;
                    MessageBox.Show(fname + " Records : " + rowcount + " exported suceessfully!", "Export DTW", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                    return false;
                }
                finally
                {
                    stream.Dispose();
                }
                return true;

            }

            return false;
        }

        #endregion

        #region " DB Related "

        public static string ConnectedtoDB()
        {
            string rs ="";
            try
            {
                int err = 0;

                if (PublicVariable.mTiCompany == null)
                {
                    PublicVariable.mTiCompany = new SAPbobsCOM.Company();
                }

                if (PublicVariable.mTiCompany.Connected)
                {
                    if (PublicVariable.mTiCompany.CompanyDB.ToString() == ConfigurationManager.AppSettings["SAPDBName"])
                    {
                        //already connected to the target company
                        rs = "";
                    }
                    else
                    {
                        PublicVariable.mTiCompany.Disconnect();
                    }
                }
                switch (ConfigurationManager.AppSettings["SAPServerType"])
                {
                    case "MSSQL2005":
                        PublicVariable.mTiCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005;
                        break;
                    case "MSSQL2008":
                        PublicVariable.mTiCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008;
                        break;
                    case "MSSQL2012":
                        PublicVariable.mTiCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;
                        break;
                    case "MSSQL2014":
                        PublicVariable.mTiCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014;
                        break;
                    case "MSSQL2016":
                        PublicVariable.mTiCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016;
                        break;
                    case "MSSQL2017":
                        PublicVariable.mTiCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2017;
                        break;
                }

                PublicVariable.mTiCompany.Server = ConfigurationManager.AppSettings["SAPServer"];
                PublicVariable.mTiCompany.CompanyDB = ConfigurationManager.AppSettings["SAPDBName"];
                PublicVariable.mTiCompany.UserName = ConfigurationManager.AppSettings["SAPUser"];
                PublicVariable.mTiCompany.Password = ConfigurationManager.AppSettings["SAPUserPsw"];
                PublicVariable.mTiCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
                err = PublicVariable.mTiCompany.Connect();

                if (err != 0)
                {
                 
                    throw new Exception(PublicVariable.mTiCompany.GetLastErrorCode() + " | " + PublicVariable.mTiCompany.GetLastErrorDescription());
                }
                else
                {
                    return "";
                }
               
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return rs;
        }

        public static void DisConnectToDB()
        {
            try
            {
                if (PublicVariable.mTiCompany != null)
                {
                    if (PublicVariable.mTiCompany.Connected)
                    {
                        if (PublicVariable.mTiCompany.CompanyDB.ToString() == ConfigurationManager.AppSettings["SAPDBName"])
                        {
                            //already connected to the target company
                            PublicVariable.mTiCompany.Disconnect();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion

        #region " Execute Query "

        public DataTable SAP_Local_RunQuery(string querystr)
        {

            mDT = new DataTable();
            try
            {
                if (Conn.State == ConnectionState.Closed)
                {
                    Conn.Open();
                }

                Cmd.CommandType = CommandType.Text;
                Cmd.Connection = Conn;
                Cmd.CommandText = querystr;
                Cmd.CommandTimeout = 0;

                SQLAdapter.SelectCommand = Cmd;


                mDT.Clear();
                SQLAdapter.Fill(mDT);

                CloseConn();


            }
            catch (SqlException sqlEx)
            {
                return new DataTable();
            }
            catch (Exception ex)
            {
                return new DataTable();
            }
            finally
            {
                if ((Conn != null) & Conn == null)
                {
                    Conn.Close();
                }
            }
            return mDT;
        }

        public String SAP_Local_RunQuery_NoResult(string querystr)
        {
            int result = 0;
            try
            {
                if (Conn.State == ConnectionState.Closed)
                {
                    Conn.Open();
                }

                Cmd.CommandType = CommandType.Text;
                Cmd.Connection = Conn;
                Cmd.CommandText = querystr;
                Cmd.CommandTimeout = 0;
                //SQLAdapter.SelectCommand = Cmd;

                result = Convert.ToInt32(Cmd.ExecuteNonQuery());

                return "";

            }
            catch (SqlException sqlEx)
            {

                return sqlEx.Message;
            }
            catch (Exception ex)
            {

                return ex.Message;
            }

            finally
            {
                if ((Conn != null) & Conn == null)
                {
                    Conn.Close();
                }
            }

        }

        public void CloseConn()
        {
            if (Conn.State == ConnectionState.Open)
            {
                Conn.Close();
            }
        }

        #endregion

        #region "Write Log"

        public void WriteFile(String mData, String filename)
        {
            mData = DateTime.Now.ToString() + mData;

            string folderpath = ConfigurationManager.AppSettings["LogFolder"];

            if (!Directory.Exists(folderpath))
            {
                System.IO.Directory.CreateDirectory(ConfigurationManager.AppSettings["LogFolder"]);
            }

            String path = Path.Combine(folderpath, filename + ".txt");

            if (!File.Exists(path))
            {
                using (StreamWriter writer = File.CreateText(path))
                {
                    writer.WriteLine(string.Format(mData));
                    writer.Close();
                }
            }
            else
            {
                using (StreamWriter writer = File.AppendText(path))
                {
                    writer.WriteLine(string.Format(mData));
                    writer.Close();
                }
            }

        }

        public void WriteLogMessageToSQL(string aLogMsg, string aIncludeError, string aLogType)
        {
            try
            {
                aLogMsg = aLogMsg.Replace("'", " ");

                string l_date = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");

                if (Conn.State == ConnectionState.Closed)
                {
                    Conn.Open();
                }
                string l_query = @"INSERT INTO [DMS_BUFFER].[dbo].[_OPS_SERVICE_LOG]([LogDate],[LogMessage],[IncludeError],[LogType])
                                            VALUES ('" + l_date + "',N'" + aLogMsg + "','" + aIncludeError + "','" + aLogType + "')";

                Cmd.CommandType = CommandType.Text;
                Cmd.Connection = Conn;
                Cmd.CommandText = l_query;
                Cmd.CommandTimeout = 0;

                int res = Convert.ToInt32(Cmd.ExecuteNonQuery());
                if (Conn.State == ConnectionState.Open)
                {
                    Conn.Close();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

        }


        #endregion

    }
}
