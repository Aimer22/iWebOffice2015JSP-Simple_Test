<%--
  Created by IntelliJ IDEA.
  User: Administrator
  Date: 2021-3-10
  Time: 9:03
  To change this template use File | Settings | File Templates.
--%>
<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%
    //项目名称 比如 /iWebOffice2015/index.jsp
    String mHttpUrlName=request.getRequestURI();
    //System.out.println("mHttpUrlName :" +mHttpUrlName);
    //得到当前页面所在目录下全名称 比如 index.jsp
    String mScriptName=request.getServletPath();
//  System.out.println("mScriptName :" +mScriptName);
    String mServerUrl="http://"+request.getServerName()+":"+request.getServerPort()+mHttpUrlName.substring(0,mHttpUrlName.lastIndexOf(mScriptName));
//  System.out.println("mServerUrl: "+mServerUrl);
    String mSessionID = request.getSession().getId();
    // System.out.println("Word.js SessionID = " + mSessionID);
%>
<html>
<head>
    <title>Title</title>
    <link rel="stylesheet" href="css/base.css">
    <link rel="stylesheet" href="css/index.css">
    <script src="js/WebOffice.js"></script>
    <script type="text/javascript">
        var WebOfficeObj = new WebOffice2015();


    </script>

    <!-- 控件加载默认调用函数  -->
    <script language="javascript" for="WebOffice2015" event="OnReady()">
        WebOfficeObj.setObj(document.getElementById('WebOffice2015'));//给2015对象赋值
        Load();//避免页面加载完，控件还没有加载情况
    </script>

    <script type="text/javascript">
        function OnReady(){
            WebOfficeObj.setObj(document.getElementById('WebOffice2015'));
            Load();
        }
    </script>
    <script type="text/javascript">

        

        function Load() {
            WebOfficeObj.ServerUrl="<%=mServerUrl%>";
            WebOfficeObj.UserName = "曹帅";
            WebOfficeObj.FileType = ".doc";
            WebOfficeObj.Author = "曹帅";
            WebOfficeObj.FileName = "test.doc";
            WebOfficeObj.AddCustomMenu();
            WebOfficeObj.AppendMenu("1","打开本地文档(&L)");
            WebOfficeObj.AppendMenu("4","打印预览(&P)");
            WebOfficeObj.AppendMenu("5","另存为")
            // WebOfficeObj.DataBase = "MYSQL";
            WebOfficeObj.obj.WebCreateProcess();
            WebOfficeObj.obj.SetEnableBtn = true;
            /*创建一个空白文档*/
            WebOfficeObj.CreateFile();

        /*创建超阅控件*/
            // WebOfficeObj.obj.CreateNew("{5a22176d-118f-4185-9653-9f98958a6df8}");
            // WebOfficeObj.obj.ActiveDocument.openFile("C:\\Users\\Administrator\\AppData\\Roaming\\SurRead\\bin\\docs\\超阅演示文档.ofd",false);


        }

        function OnCommand(ID, Caption, bCancel){
            switch(ID){
                case 1:OpenLocalFile();break;//打开本地文件
                case 4:WebOfficeObj.PrintPreview();break;//启用
                case 5:WebOfficeObj.setShowDialog("E:\\IntelliJ IDEA 2021.1.3\\workspace\\iWebOffice2015JSP-Simple_Test\\out\\artifacts\\iWebOffice2015JSP_Simple_Test_Web_exploded\\document\\测试文档.docx"); break
                default:;return;
            }
        }


        function OpenFile() {
            if(WebOfficeObj.WebOpen()){
                WebOfficeObj.Alert("打开文档成功");
            }else{
                WebOfficeObj.Alert("打开失败55");
            }
        }

        function OpenLocalFile(){
           try{
                   var filePath =  WebOfficeObj.WebOpenLocal();
                   WebOfficeObj.Alert(filePath);
               WebOfficeObj.WebSetRevision(false,true,false,false);
           }catch (e){
               WebOfficeObj.Alert("打开失败");
           }
        }
        
        function OpenFileByURL() {
            try{

                //WebOfficeObj.WebUrl = "http://www.kinggrid.com:8080/iWebOffice2015/OfficeServer";
                WebOfficeObj.ServerUrl = "http://www.kinggrid.com:8080/iWebOffice2015";
                WebOfficeObj.FileName = "1615364833464.doc";
                WebOfficeObj.WebOpen();
            }catch (e) {
                WebOfficeObj.Alert(e.printStackTrace());
            }
        }

        function CreateBookMarks() {
            try{
                WebOfficeObj.WebAddBookMarks("content","在此处添加内容");
                WebOfficeObj.VBAInsertFile("content","C:\\Users\\99685\\Desktop\\常见问题标题整理-曹帅.docx");

            }catch (e) {
                WebOfficeObj.Alert("调用接口失败")
            }
        }
        
        function  WebUseTemplate() {
            try {
                //先从服务器中加载文档
                WebOfficeObj.Template = "公函---原格式.docx";
                WebOfficeObj.WebUseTemplate();
                WebOfficeObj.VBAInsertFile("content","");
                console.log("模板添加成功！！");

            }catch (e) {
               WebOfficeObj.Alert("调用接口失败")
            }
        }
        //保存文档至服务器
        function WebSave() {
            try {
                if(WebOfficeObj.WebSave()){
                    WebOfficeObj.Alert("保存文档成功");
                }
            }catch (e) {
                WebOfficeObj.Alert("调用接口失败")
            }
        }

        //保存文档到数据库
        function WebSaveToDB(){
            try {
                WebOfficeObj.DataBase = "MYSQL";
                if(WebOfficeObj.WebSave()){
                    WebOfficeObj.Alert("保存文档成功");
                }
            }catch (e) {
                WebOfficeObj.Alert("调用接口失败")
            }
        }
        //添加水印
        function AddWaterMarks() {
            try {
                WebOfficeObj.AddWaterMark("金格科技");
            }catch (e) {
                WebOfficeObj.Alert("调用接口失败")
            }
        }
        //清楚当前文档的痕迹
        function ClearRevision() {
            try {
                if(WebOfficeObj.ClearRevisions()){
                    WebOfficeObj.Alert("成功接受文档的痕迹");
                }else{
                    WebOfficeObj.Alert("接受痕迹失败");
                }
            }catch (e) {
                WebOfficeObj.Alert("调用接口失败");
            }
        }

        //测试打印
        function WebPrint(){
            try {
                WebOfficeObj.WebOpenPrint();
            }catch (e) {
                WebOfficeObj.Alert("调用接口失败");
            }
        }

        //测试文档添加水印
        function AddWaterMarks(){
            try {
                WebOfficeObj.AddWaterMark("金格科技");
            }catch (e) {
                WebOfficeObj.Alert("调用接口失败");
            }
        }

        //测试插入图片水印
        function AddImgWaterMarks(){
            try {
                WebOfficeObj.AddGraphicWaterMark("金格科技");
            }catch (e) {
                WebOfficeObj.Alert("调用接口失败");
            }
        }

        function WebSavePDF(){
            try{
                WebOfficeObj.RecordID = "1111";
                WebOfficeObj.WebSavePDF();
            }catch (e) {
                WebOfficeObj.Alert("调用接口失败");
            }
        }

        function WebSetProtect(){
            try{
                WebOfficeObj.WebSetProtect(true,"");
                WebOfficeObj.Alert("文档已保护");
            }catch (e){
                WebOfficeObj.Alert("调用接口失败");
            }
        }

        function CreateNewFile(){
            try{
                WebOfficeObj.FileName = "ShyeChou";
                WebOfficeObj.CreateFile();

            }catch (e){
                WebOfficeObj.Alert("调用接口失败");
            }
        }

        function INetSetCookie(){
            try{
                WebOfficeObj.INetSetCookie("<%=mServerUrl%>","<%=mSessionID%>");

            }catch (e){
                WebOfficeObj.Alert("调用接口失败");
            }
        }

        function WebDeleteContentInBookMarks(){
            try{
                WebOfficeObj.WebDeleteContentInBookMarks("Book1","Book2");
                WebOfficeObj.Alert("删除成功！");

            }catch (e){
                WebOfficeObj.Alert("调用接口失败");
            }
        }

        function BookMarksPosition(){
            try{
                WebOfficeObj.WebFindBookMarks("Book2");

            }catch (e){
                WebOfficeObj.Alert("调用接口失败");
            }
        }

        //动态添加书签
        function AddBookMarks(){
            try{
                WebOfficeObj.WebAddBookMarks("Content");

            }catch (e){
                WebOfficeObj.Alert("调用接口失败");
            }
        }
    </script>
</head>
<body onload="">
<header>
    <div class="logo">
        <img src="http://www.kinggrid.com/img/logo.jpg" alt="">
    </div>
</header>
<main class="main w clearfix" >
    <div class="nav">
        <ul class="nav-menu">
            <li><a href="#" onclick="OpenFile()">打开文档</a></li>
            <li><a href="#" onclick="OpenLocalFile()">打开本地文档</a></li>
            <li><a href="#" onclick="WebSave()">保存文档</a></li>
            <li><a href="#" onclick="WebSaveToDB()">保存文档至数据库</a></li>
            <li><a href="#" onclick="WebUseTemplate()">套红应用</a></li>
            <li><a href="#" onclick="CreateBookMarks()">书签填充</a></li>
            <li><a href="#" onclick="AddWaterMarks()">插入水印</a></li>
            <li><a href="#" onclick="OpenFileByURL();">打开URL文档</a></li>
            <li><a href="#" onclick="ClearRevision()">清理痕迹</a></li>
            <li><a href="#" onclick="WebOfficeObj.WebShow(true)">显示痕迹</a></li>
            <li><a href="#" onclick="WebOfficeObj.WebShow(false)">隐藏痕迹</a></li>
            <li><a href="#" onclick="WebPrint()">打印</a></li>


            <li><a href="#" onclick="AddWaterMarks()">添加水印</a></li>
            <li><a href="#" onclick="AddImgWaterMarks()">添加图片水印</a></li>
            <li><a href="#" onclick="WebOfficeObj.AddWaterMark1()">添加文字水印（兼容）</a></li>
            <li><a href="#" onclick="WebSavePDF()">保存PDF到服务器</a></li>
            <li><a href="#" onclick="WebOfficeObj.ParagraphSettings()">设置段落</a></li>
            <li><a href="#" onclick="WebSetProtect()">保护文档（VBA）</a></li>
            <li><a href="#" onclick="WebOfficeObj.WebSaveLocal()">保存文档到本地</a></li>
            <li><a href="#" onclick="WebOfficeObj.ShowStatusBar(true)">显示状态栏</a></li>
            <li><a href="#" onclick="WebOfficeObj.ShowStatusBar(false)">隐藏状态栏</a></li>
            <li><a href="#" onclick="CreateNewFile()">创建新文档（并且命名）</a></li>
            <li><a href="#" onclick="WebOfficeObj.ParagraphSettings()">设置段落</a></li>
            <li><a href="#" onclick="WebOfficeObj.INetSetCookie()">设置Cookie</a></li>
            <li><a href="#" onclick="AddBookMarks()">动态添加书签</a></li>
            <li><a href="#" onclick="WebDeleteContentInBookMarks()">删除两个标签之间的内容</a></li>
            <li><a href="#" onclick="WebOfficeObj.WebSetRibbonUIXML()">禁用审阅</a></li>
            <li><a href="#" onclick="WebOfficeObj.getRevision()">获取文档痕迹</a></li>


        </ul>
        <div class="weboffice" id="weboffice"><script src="js/iWebOffice2015.js"></script></div>
        <span id = "OfficeDiv">控件加载成功</span>
    </div>
</main>
<footer class="footer">
    <p>赣ICP备14004497号-1 2003-2021 江西金格科技股份有限公司 版权所有</p>
</footer>

</body>
</html>
