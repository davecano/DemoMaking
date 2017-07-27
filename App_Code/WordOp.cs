using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
//引用Interop.Microsoft.Office.Interop.Word.dll
using Microsoft.Office.Interop.Word;
using ICSharpCode.SharpZipLib.Zip;
using ICSharpCode.SharpZipLib.Checksums;
//using ICSharpCode.SharpZipLib.Core;
using System.IO;
// <summary>
//WordOp
// </summary>
public class WordOp
{
    public WordOp()
    {
        //TODO: 在此处添加构造函数逻辑
    }
private ApplicationClass WordApp;
    private Document WordDoc;
    private static bool isOpened = false;//判断word模版是否被占用

    public void SaveAs(string strFname, bool isReplace)
    {
        if (isReplace && File.Exists(strFname))
        { File.Delete(strFname); }
        object missing = Type.Missing;
        object fileName = strFname;
        WordDoc.SaveAs(ref fileName, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
    }
//定义一个Word.Application 对象

public void activeWordApp()
    { WordApp = new ApplicationClass(); }
    public void Quit()
    {
        object missing = System.Reflection.Missing.Value;
        WordApp.Application.Quit(ref missing, ref missing, ref missing);
        isOpened = false;
    }
    //按照先前设计好的模版新建Word文件
    public void OpenTempelte(string strTemppath)
    {
        object Missing = Type.Missing;
        //object Missing = System.Reflection.Missing.Value;
        activeWordApp();
        WordApp.Visible = false;
        object oTemplate = (object)strTemppath;
        try
        {
            WordDoc = WordApp.Documents.Add(ref oTemplate, ref Missing, ref Missing, ref Missing);
            isOpened = true;
            WordDoc.Activate();
        }
        catch (Exception Ex)
        {
            Quit();
            isOpened = false;
            throw new Exception(Ex.Message);
        }
    }
    public void FillLable(string LabelId, string Content)
    { //打开Word模版
      // OpenTempelte(tempName); //对LabelId的标签进行填充内容Content,即函件题目项
        object bkmC = LabelId;
        if (WordApp.ActiveDocument.Bookmarks.Exists(LabelId) == true)
        {
            if (LabelId != "PIC")//判断是否是显示照片的书签
            {
                WordApp.ActiveDocument.Bookmarks.get_Item(ref bkmC).Select();
                WordApp.Selection.TypeText(Content);
            }
            else
            {
                try
                {
                    object missing = System.Reflection.Missing.Value;
                    InlineShape li = WordApp.ActiveDocument.Bookmarks.get_Item(ref bkmC).Range.InlineShapes.AddPicture(Content,
                    ref missing, ref missing, ref missing);
                    li.Width = 85;//设置照片的宽
                    li.Height = 100;
                } //设置照片的高
                catch { }
            }
        }
    }
    public static void PackFiles(string filename, string directory)
    {
        try
        {
            FastZip fz = new FastZip();
            fz.CreateEmptyDirectories = true;
            fz.CreateZip(filename, directory, true, "");
            fz = null;
        }
        catch (Exception)
        {
            throw;
        }
    }

}
