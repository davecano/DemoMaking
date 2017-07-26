using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class form1 : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        //if (!IsPostBack) {
        //HyperLink1.Visible = true;
        //HyperLink1.NavigateUrl = Server.MapPath("download") + "/data.rar";
    
        //}
       
    }

    protected void Button1_Click(object sender, EventArgs e)
    {
        modeluser mu = new modeluser();
        mu.UserID = "04130215";
        mu.UserName = "陈旋";
        mu.UserPsw = "123";
        mu.Content = "this is a paragraph";
        modeluser mu1 = new modeluser();
        mu1.UserID = "04130216";
        mu1.UserName = "davecano";
        mu1.UserPsw = "123";
        mu1.Content = "this is another paragraph";
        List<modeluser> li = new List<modeluser>();
        li.Add(mu);
        li.Add(mu1);
        WordOp wop = new WordOp();
        foreach (modeluser m in li) {
       string path = Server.MapPath("temptate");
        string templatePath = path + "//temptatetest.docx";
        wop.OpenTempelte(templatePath);
        wop.FillLable("UserID", m.UserID);
        wop.FillLable("UserName",m.UserPsw);
        wop.FillLable("UserPsw", m.UserPsw);
        wop.FillLable("Content",m.Content);
        wop.SaveAs(Server.MapPath("targetword") + "//"+m.UserID+"doc", true);
        wop.Quit();
        }
        wop.CreateRar(Server.MapPath("targetword"), Server.MapPath("download") + "/data.rar");
        Response.Write("ok,now press the download button...");
        LinkButton1.Visible = true;
        //HyperLink1.Visible = true;
        //HyperLink1.NavigateUrl = Server.MapPath("download") + "/data.rar";
    }



    protected void LinkButton1_Click(object sender, EventArgs e)
    {
        //服务器文件路径
        string strFilePath = Server.MapPath("download") + "/data.rar";
        FileInfo fileInfo = new FileInfo(strFilePath);
        Response.Clear();
        Response.AddHeader("content-disposition", "attachment;filename=" + Server.UrlEncode(fileInfo.Name.ToString()));
        Response.AddHeader("content-length", fileInfo.Length.ToString());
        Response.ContentType = "application/octet-stream";
        Response.ContentEncoding = System.Text.Encoding.Default;
        Response.WriteFile(strFilePath);
    }
}