using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using System.Collections;

/// <summary>
/// Class1 的摘要描述
/// </summary>
public class sapReportPrms
{
            string cs = "Data Source=192.168.0.9;Initial Catalog=rptsvrDB ;User ID=rptsvrDB ;Password=adminrp";
            string qs = "select * from rptsvrDB.dbo.sapReportPrms where(1=1) AND sapReportPrms.rptNm ='CO11N'";
          
            String D_connIP,D_connUser,  D_connPwd,  D_rptNm, D_status, D_connClient,  D_connLanguage;
            string []ALL;


            public string[] SQL()

            {
            try
              {

                  //1.引用SqlConnection物件連接資料庫
                  SqlConnection cn = new SqlConnection(cs);
                  //2.開啟資料庫
                  cn.Open();
                  //3.引用SqlCommand物件
                  SqlCommand command = new SqlCommand(qs, cn);
                  //4.搭配SqlCommand物件使用SqlDataReader
                  SqlDataReader dr = command.ExecuteReader();
                  while ((dr.Read()))
                  {
                      int intIP = dr.GetOrdinal("connIP");
                      int intUser = dr.GetOrdinal("connUser");
                      int intPwd = dr.GetOrdinal("connPwd");
                      int intrptNm = dr.GetOrdinal("rptNm");
                      int intstatus = dr.GetOrdinal("status");
                      int intclient = dr.GetOrdinal("connClient");
                      int intlanguage = dr.GetOrdinal("connLanguage");
                      D_connIP = dr[intIP].ToString();
                      D_connUser = dr[intUser].ToString();
                      D_connPwd =dr[intPwd].ToString();
                      D_rptNm = dr[intrptNm].ToString();
                      D_status = dr[intstatus].ToString();
                      D_connClient = dr[intclient].ToString();
                      D_connLanguage = dr[intlanguage].ToString();

                       ALL = new string[7] { D_connIP, D_connUser, D_connPwd, D_rptNm, D_status,
                          D_connClient, D_connLanguage };
                  }
                  cn.Close();
                return ALL;
             }
                finally 
             
               {

          }
    
         } 
  
}
