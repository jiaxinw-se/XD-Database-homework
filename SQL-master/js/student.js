  // 教师查询所有学籍信息
function searchAllMessage(){
    var conn = new ActiveXObject("ADODB.Connection");
    conn.Open("DBQ=e://dataBase/Studat1.mdb;DRIVER={Microsoft Access Driver (*.mdb)};");
    var rs = new ActiveXObject("ADODB.Recordset");
    var sql = "select * from Student";
    rs.open(sql, conn);

    var studentTable = document.getElementById("studentTable"); 
    studentTable.style.visibility = 'visible';   //显示学籍表

    var length= studentTable.rows.length;               //获得Table下的行数 
    if(length!=1){                                      //如果有行，则清空 
        for(var i=length-1;i>=1;i--)  
            {  
                studentTable.deleteRow(i);     
                }  
    } 

    while(!rs.EOF)
    { 
        var trow = getDataRow(rs);
        studentTable.appendChild(trow);

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
        delBtn.setAttribute("onclick","delRow(this)");
        delCell.appendChild(delBtn);
        trow.appendChild(delCell);

        rs.moveNext();
    }
    rs.close();
    rs = null;
    conn.close();
    conn = null;
}

// 增加
//添加学籍信息
function addUser(id,stuName,stuSex,stuAge,stuGrade){
    //用 JavaScript 写服务器端连接数据库
    var conn = new ActiveXObject("ADODB.Connection");
    conn.Open("DBQ=d://dataBase/Studat1.mdb;DRIVER={Microsoft Access Driver (*.mdb)};");
    var sql="insert into Student(Student_id,Student_name,Student_sex,Student_age,Student_grade) values("+id+",'"+stuName+"','"+stuSex+"','"+stuAge+"','"+stuGrade+"')";
    try{
        conn.execute(sql);
        alert("添加成功");
        searchAllMessage();
    }
    catch(e){
        alert("添加失败~~~"+e.description);
    }
    conn.close();
}


//把数据库中的一行显示到页面上
function getDataRow(rs){
    var row = document.createElement("tr");//创建行
    var idCell = document.createElement("td");//创建id列
    idCell.innerHTML = rs.Fields("Student_id");//填充id
    row.setAttribute("id",idCell.innerHTML);
    row.appendChild(idCell);
    var nameCell = document.createElement("td");
    nameCell.innerHTML = rs.Fields("Student_name");
    row.appendChild(nameCell);
    var sexCell = document.createElement("td");//创建性别列
    sexCell.innerHTML = rs.Fields("Student_sex");//填充性别
    row.appendChild(sexCell);
    var ageCell = document.createElement("td");
    ageCell.innerHTML = rs.Fields("Student_age");
    row.appendChild(ageCell);
    var gradeCell = document.createElement("td");
    gradeCell.innerHTML = rs.Fields("Student_grade");
    row.appendChild(gradeCell);
    return row;
}

//删除前端数据操作
function delRow(obj){
    var stuName = document.getElementById(obj.parentNode.parentNode.id).children[1].innerHTML;//获取被删除学生的名字
    if(confirm("确定删除"+stuName+"同学嘛？")){ 
        delStu(obj.parentNode.parentNode.id);
        //找到按钮所在行的节点，然后删掉这一行 
        obj.parentNode.parentNode.parentNode.removeChild(obj.parentNode.parentNode); 
        //btnDel - td - tr - tbody - 删除(tr) 
        //刷新网页还原。
    } 
}

//删除数据库数据操作
function delStu(id){
    var conn = new ActiveXObject("ADODB.Connection");
    conn.Open("DBQ=d://dataBase/Studat1.mdb;DRIVER={Microsoft Access Driver (*.mdb)};");
    var sql = "delete from Student where Student_id = "+"'"+id+"'";
    conn.execute(sql);
    conn.close();
    conn = null;
}

//修改信息
function updateUser(){

}

//学生查看自己的学籍信息
function searchMyMessage(myID){
    var conn = new ActiveXObject("ADODB.Connection");
    conn.Open("DBQ=e://dataBase/Studat1.mdb;DRIVER={Microsoft Access Driver (*.mdb)};");
    var rs = new ActiveXObject("ADODB.Recordset");
    var sql="select * from Student where Student_id ="+"'"+myID+"'";        //根据学生ID查询学籍信息的SQL代码
    rs.open(sql, conn);

    var studentTable = document.getElementById("studentTable");
    var scoreTable = document.getElementById("scoreTable");
    studentTable.style.visibility = 'visible';
    scoreTable.style.visibility = 'hidden';
    var length= studentTable.rows.length;               //获得Table下的行数 
    if(length!=1){                                      //如果有行，则清空 
        for(var i=length-1;i>=1;i--){  
                studentTable.deleteRow(i);     
        }  
    }  


    while(!rs.EOF)
    { 
        var trow = getDataRow(rs);
        studentTable.appendChild(trow);
        rs.moveNext();
    }

    // var boxHeight = scoreTable.offsetHeight;  //获取表格高度
    // var boxWidth = scoreTable.offsetWidth; //获取表格宽度
    var MyTable = document.getElementById("MyTable");
    MyTable.style.height= "108px";
    MyTable.style.width = "732px";
    rs.close();
    rs = null;
    conn.close();
    conn = null;
}

//修改学生学籍信息
function changeStuMessage(){
    alert("修改成功");
}
