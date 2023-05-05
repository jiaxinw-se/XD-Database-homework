function checkID(obj,length){
    if(obj.length !=length){
        obj.setCustomValidity("请输入10位学号");
    }
    else if(obj.length == 0){
        obj.setCustomValidity("学号不能为空");
    }
    else{
        obj.setCustomValidity("");
    }
}