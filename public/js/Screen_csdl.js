function fn_Table01_SQL_Show(){
    socket.emit("msg_SQL_Show", "true");
    socket.on('SQL_Show',function(data){
        fn_table_01(data);
    });
}

function fn_Table01_SQL_KHTT(){
    socket.emit("msg_SQL_KHTT", "true");
    socket.on('SQL_KHTT',function(data){
        fn_table_01(data);
    });
}

// Chương trình con hiển thị SQL ra bảng
function fn_table_01(data){
    if(data){
        $("#table_01 tbody").empty();
        var len = data.length;
        var txt = "<tbody>";
        if(len > 0){
            for(var i=0;i<len;i++){
                    txt += "<tr><td style='text-align: center;'>"+data[i].ID
                        +"</td><td style='text-align: center;'>"+data[i].Num
                        +"</td><td style='text-align: center;'>"+data[i].Color
                        +"</td><td style='text-align: center;'>"+data[i].Time_In
                        +"</td><td style='text-align: center;'>"+data[i].Time_Out
                        +"</td><td style='text-align: center;'>"+data[i].Position
                        +"</td><td style='text-align: center;'>"+data[i].Price
                        +"</td><td style='text-align: center;'>"+data[i].Member
                        +"</td></tr>";
                    }
            if(txt != ""){
            txt +="</tbody>";
            $("#table_01").append(txt);
            }
        }
    }
}

// Tìm kiếm SQL theo khoảng thời gian
function fn_SQL_By_Time()
{
    var startTime = document.getElementById('dtpk_Search_Start').value;
    var endTime = document.getElementById('dtpk_Search_End').value;

    // Kiểm tra nếu ô nhập liệu thời gian không có giá trị
    if (startTime === '' || endTime === '') {
        alert('Vui lòng nhập đầy đủ thời gian để tìm kiếm.');
        return; // Dừng hàm nếu không đủ thời gian
    }

    // Đóng gói giá trị và thời gian vào một mảng
    var val = [startTime, endTime];

    socket.emit('msg_SQL_ByTime', val);
    socket.on('SQL_ByTime', function(data){
        fn_table_01(data); // Show sdata
    });
}

function fn_SQL_By_Time_In()
{
    var startTime = document.getElementById('dtpk_Search_Start').value;
    var endTime = document.getElementById('dtpk_Search_End').value;

    // Kiểm tra nếu ô nhập liệu thời gian không có giá trị
    if (startTime === '' || endTime === '') {
        alert('Vui lòng nhập đầy đủ thời gian để tìm kiếm.');
        return; // Dừng hàm nếu không đủ thời gian
    }

    // Đóng gói giá trị và thời gian vào một mảng
    var val = [startTime, endTime];

    socket.emit('msg_SQL_ByTime_In', val);
    socket.on('SQL_ByTime_In', function(data){
        fn_table_01(data); // Show sdata
    });
}

function fn_SQL_By_Time_Out()
{
    var startTime = document.getElementById('dtpk_Search_Start').value;
    var endTime = document.getElementById('dtpk_Search_End').value;

    // Kiểm tra nếu ô nhập liệu thời gian không có giá trị
    if (startTime === '' || endTime === '') {
        alert('Vui lòng nhập đầy đủ thời gian để tìm kiếm.');
        return; // Dừng hàm nếu không đủ thời gian
    }

    // Đóng gói giá trị và thời gian vào một mảng
    var val = [startTime, endTime];

    socket.emit('msg_SQL_ByTime_Out', val);
    socket.on('SQL_ByTime_Out', function(data){
        fn_table_01(data); // Show sdata
    });
}

function fn_SQL_By_Num_Car()
{
    var Num_car = document.getElementById('txt_ValueToSearch_Num_Car').value;

    // Kiểm tra nếu ô nhập liệu thời gian không có giá trị
    if (Num_car === '') {
        alert('Vui lòng nhập đầy đủ để tìm kiếm.');
        return; // Dừng hàm nếu không đủ thời gian
    }

    // Đóng gói giá trị và thời gian vào một mảng
    var val = Num_car;

    socket.emit('msg_SQL_ByNum_Car', val);
    socket.on('SQL_ByNum_Car', function(data){
        fn_table_01(data); // Show sdata
    });
}

function fn_SQL_By_Color_Car()
{
    var Color_car = document.getElementById('txt_ValueToSearch_Color_Car').value;

    // Kiểm tra nếu ô nhập liệu thời gian không có giá trị
    if (Color_car === '') {
        alert('Vui lòng nhập đầy đủ để tìm kiếm.');
        return; // Dừng hàm nếu không đủ thời gian
    }

    // Đóng gói giá trị và thời gian vào một mảng
    var val = Color_car;

    socket.emit('msg_SQL_ByColor_Car', val);
    socket.on('SQL_ByColor_Car', function(data){
        fn_table_01(data); // Show sdata
    });
}

function fn_SQL_By_Position_Car()
{
    var Position_car = document.getElementById('txt_ValueToSearch_Position_Car').value;

    // Kiểm tra nếu ô nhập liệu thời gian không có giá trị
    if (Position_car === '') {
        alert('Vui lòng nhập đầy đủ để tìm kiếm.');
        return; // Dừng hàm nếu không đủ thời gian
    }

    // Đóng gói giá trị và thời gian vào một mảng
    var val = Position_car;

    socket.emit('msg_SQL_ByPosition_Car', val);
    socket.on('SQL_ByPosition_Car', function(data){
        fn_table_01(data); // Show sdata
    });
}
