// 学生查看自己的各科成绩
function searchMyScore(myID){
    var conn = new ActiveXObject("ADODB.Connection");
    conn.Open("DBQ=e://dataBase/Studat1.mdb;DRIVER={Microsoft Access Driver (*.mdb)};");
    var rs = new ActiveXObject("ADODB.Recordset");
    var sql = "select  * from Score_view where Student_id = "+"'"+myID+"'";   //获取学生各科成绩的视图  视图格式：课程ID 课程名称 学生姓名 学生学号 学生成绩 取得学分       
    rs.open(sql,conn);

    var studentTable = document.getElementById("studentTable"); 
    var scoreTable = document.getElementById("scoreTable");
    studentTable.style.visibility = 'hidden';       //隐藏学籍表
    scoreTable.style.visibility = 'visible';        //显示选课表


    var length= scoreTable.rows.length;               //获得Table下的行数 
    if(length!=1){                                      //如果有行，则清空 
        for(var i=length-1;i>=1;i--)  
            {  
                scoreTable.deleteRow(i);     
                }  
    }
    
    while(!rs.EOF)
    {
        var trow = getScoreRow(rs);
        scoreTable.appendChild(trow);
        rs.moveNext();
    }

    var boxHeight = scoreTable.offsetHeight;  //获取表格高度
    var boxWidth = scoreTable.offsetWidth; //获取表格宽度
    var MyTable = document.getElementById("MyTable");
    MyTable.style.height = boxHeight+"px";
    MyTable.style.width = "732px";

    rs.close();
    rs = null;
    conn.close();
    conn = null;
}



// 获取学生成绩表中的每行数据
function getScoreRow(rs){
    var row = document.createElement("tr");//创建行
    var classIDCell = document.createElement("td");//创建课程id列
    classIDCell.innerHTML = rs.Fields("Class_id"); //填充课程id
    var classNameCell = document.createElement("td");//创建课程名称列
    classNameCell.innerHTML = rs.Fields("Class_name"); //填充课程名称
    var nameCell = document.createElement("td");    //学生名称列
    nameCell.innerHTML = rs.Fields("Student_name"); //填充学生名称
    var idCell = document.createElement("td");//创建学生id列
    idCell.innerHTML = rs.Fields("Student_id");//填充id
    var studentScoreCell = document.createElement("td");    //成绩
    studentScoreCell.innerHTML = rs.Fields("Student_score");
    var classCountCell = document.createElement("td");   //学分
    classCountCell.innerHTML = rs.Fields("Class_count");
    row.appendChild(classIDCell);
    row.appendChild(classNameCell);
    row.appendChild(idCell);
    row.appendChild(nameCell);
    row.appendChild(studentScoreCell);
    row.appendChild(classCountCell);

    return row;
}

//登记学生成绩
function addScore(inputID,inputClassID,inputScore){
    var conn = new ActiveXObject("ADODB.Connection");
    conn.Open("DBQ=e://dataBase/Studat1.mdb;DRIVER={Microsoft Access Driver (*.mdb)};");
    var sql="insert into Score(Student_id,Class_id,Student_score) values("+inputID+",'"+inputClassID+"',"+inputScore+")";
    try{
        conn.execute(sql);
        alert("添加成功");
        searchAllScore();
    }
    catch(e){
        alert("添加失败~~~"+e.description);
    }
    conn.close();
}


//登记新的课程
function addClass(classID,className,classCount){
    var conn = new ActiveXObject("ADODB.Connection");
    conn.Open("DBQ=e://dataBase/Studat1.mdb;DRIVER={Microsoft Access Driver (*.mdb)};");
    var sql="insert into Class(Class_id,Class_name,Class_count) values("+classID+",'"+className+"','"+classCount+"')";
    try{
        conn.execute(sql);
        alert("添加成功");
        searchAllClass();
    }
    catch(e){
        alert("添加失败~~~"+e.description);
    }
    conn.close();
}

//查看所有学生的成绩
function searchAllScore(choice){
    var conn = new ActiveXObject("ADODB.Connection");
    conn.Open("DBQ=e://dataBase/Studat1.mdb;DRIVER={Microsoft Access Driver (*.mdb)};");
    var rs = new ActiveXObject("ADODB.Recordset");
    var sql;
    if(choice==1){
        sql = "select * from Score_view";   //获取学生各科成绩的视图  视图格式：课程ID 课程名称 学生姓名 学生学号 学生成绩 取得学分       
    }
    else if(choice==2){
        sql = "select * from Score_view where Student_score < 60"
    }
    else if(choice==3){
        sql = "select * from Score_view order by Student_score DESC"
    } 
    rs.open(sql,conn);
    var allScoreTable = document.getElementById("allScoreTable");


    var length= allScoreTable.rows.length;               //获得Table下的行数 
    if(length!=1){                                      //如果有行，则清空 
        for(var i=length-1;i>=1;i--)  
            {  
                allScoreTable.deleteRow(i);     
                }  
    }
    
    while(!rs.EOF)                                    //把数据库中获取的数据逐行放入前端
    {
        var trow = getScoreRow(rs);

        var delCell = document.createElement("td"); //创建删除列
        var delBtn = document.createElement("input"); //创建删除按钮
        var changeBtn = document.createElement("input"); //创建修改按钮
        
        changeBtn.setAttribute("type","button");
        changeBtn.setAttribute("class","delBtn");
        changeBtn.setAttribute("value","修改");
        changeBtn.setAttribute("onclick","changeStuMessage(this)");
        delCell.appendChild(changeBtn);

        delBtn.setAttribute("type","button");
        delBtn.setAttribute("class","delBtn");
        delBtn.setAttribute("value","删除");
        delCell.appendChild(delBtn);
        trow.appendChild(delCell);

        allScoreTable.appendChild(trow);
        rs.moveNext();
    }

    rs.close();
    rs = null;
    conn.close();
    conn = null;
}

//查看所有课程列表
function searchAllClass(){
    var conn = new ActiveXObject("ADODB.Connection");
    conn.Open("DBQ=e://dataBase/Studat1.mdb;DRIVER={Microsoft Access Driver (*.mdb)};");
    var rs = new ActiveXObject("ADODB.Recordset");
    var sql = "select  * from Class";    
    rs.open(sql,conn);

    var classTable = document.getElementById("classTable"); 



    var length= classTable.rows.length;               //获得Table下的行数 
    if(length!=1){                                      //如果有行，则清空 
        for(var i=length-1;i>=1;i--)  
            {  
                classTable.deleteRow(i);     
                }  
    }
    
    while(!rs.EOF)
    {
        var trow = getClassRow(rs);
        var delCell = document.createElement("td"); //创建删除列
        var delBtn = document.createElement("input"); //创建删除按钮
        delBtn.setAttribute("type","button");
        delBtn.setAttribute("class","delBtn");
        delBtn.setAttribute("value","删除");
        delCell.appendChild(delBtn);
        trow.appendChild(delCell);
        classTable.appendChild(trow);
        rs.moveNext();
    }

    rs.close();
    rs = null;
    conn.close();
    conn = null;
}


//获取每行课程数据
function getClassRow(rs){
    var row = document.createElement("tr");//创建行
    var classIDCell = document.createElement("td");//创建课程id列
    classIDCell.innerHTML = rs.Fields("Class_id"); //填充课程id
    var classNameCell = document.createElement("td");//创建课程名称列
    classNameCell.innerHTML = rs.Fields("Class_name"); //填充课程名称
    var classCountCell = document.createElement("td");   //学分
    classCountCell.innerHTML = rs.Fields("Class_count");
    row.appendChild(classIDCell);
    row.appendChild(classNameCell);
    row.appendChild(classCountCell);

    return row;
}