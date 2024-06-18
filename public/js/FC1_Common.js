// Hàm chức năng chuyển trang
function fn_ScreenChange(scr_1, scr_2, scr_3, scr_4,scr_5,scr_6)
{
    document.getElementById(scr_1).style.visibility = 'visible'; // Hiển thi trang được chọn
    document.getElementById(scr_2).style.visibility = 'hidden'; // Ấn trang 1
    document.getElementById(scr_3).style.visibility = 'hidden'; // An trang 2
    document.getElementById(scr_4).style.visibility = 'hidden';
    document.getElementById(scr_5).style.visibility = 'hidden';
    document.getElementById(scr_6).style.visibility = 'hidden';
}

//////////////// CÁC KHỐI CHƯƠNG TRÌNH CON ////////////// -->
 // Chương trình con đọc dữ liệu lên IO Field
function fn_IOFieldDataShow(tag, IOField, tofix){
    socket.on(tag,function(data){
        if(tofix == 0){
            document.getElementById(IOField).value = data;
        } else{
        document.getElementById(IOField).value = data.toFixed(tofix);
        }
    });
}

function fn_IOFieldDataShow3(tag, IOField){
    var tzset = (new Date()).getTimezoneOffset() * 60000;
    socket.on(tag,function(data){
        let dateh = Date.parse(data);
        let datej = new Date(dateh);
        timeh = (new Date(datej - tzset)).toISOString().slice(0, -1).replace("T"," ").replace(".000"," ");
        document.getElementById(IOField).value = timeh;
    });
}

function fn_IOFieldDataShow2(tag, IOField){
    socket.on(tag,function(data){
        if(data == 0){
            document.getElementById(IOField).style.backgroundColor = "gray";
        } else if(data == 1){
            document.getElementById(IOField).style.backgroundColor = "yellow";
        }
        else{
            document.getElementById(IOField).style.backgroundColor = "gray";
        }
    });
}

var myVar = setInterval(myTimer, 100);
function myTimer() {
    socket.emit("Client-send-data", "Request data client");
}

function fn_SymbolStatus(ObjectID, SymName, Tag){
    var imglink_0 = "images/symbol/" + SymName + "_1.png"; // Trạng thái tag = 0
    var imglink_1 = "images/symbol/" + SymName + "_2.png"; // Trạng thái tag = 1
    socket.on(Tag,function(data){
        if(data == 0){
            document.getElementById(ObjectID).src = imglink_0;
        }
        else if(data == 1){
            document.getElementById(ObjectID).src = imglink_1;
        }
        else{
            document.getElementById(ObjectID).src = imglink_0;
        }
    });
}

function fn_DataEdit(button1, button2)
{
    document.getElementById(button1).style.zIndex='1';  // Hiển nút 1
    document.getElementById(button2).style.zIndex='0';  // Ẩn nút 2
}


///// CHƯƠNG TRÌNH CON NÚT NHẤN SỬA //////
// Tạo 1 tag tạm báo đang sửa dữ liệu
var S2_data_edditting = false;
function fn_S2_EditBtt(){
    // Cho hiển thị nút nhấn lưu
    fn_DataEdit('btt_S2_Save','btt_S2_Edit');
    // Cho tag báo đang sửa dữ liệu lên giá trị true
    S2_data_edditting = true; 
    // Kích hoạt chức năng sửa của các IO Field
    document.getElementById("gocquay").disabled = false; // Tag Real
}
///// CHƯƠNG TRÌNH CON NÚT NHẤN LƯU //////
function fn_S2_SaveBtt(){
// Cho hiển thị nút nhấn sửa
fn_DataEdit('btt_S2_Edit','btt_S2_Save');
    // Cho tag đang sửa dữ liệu về 0
    S2_data_edditting = false; 
                        // Gửi dữ liệu cần sửa xuống PLC
    var data_edit_array = [document.getElementById('gocquay').value
                          ];
    socket.emit('cmd_S2_Edit_Data', data_edit_array);
    alert('Dữ liệu đã được lưu!');
    // Vô hiệu hoá chức năng sửa của các IO Field
    document.getElementById("gocquay").disabled = true;    // Tag Real
}
 
// Chương trình con đọc dữ liệu lên IO Field
function fn_S2_IOField_IO(tag, IOField, tofix)
{
    socket.on(tag, function(data){
        if (tofix == 0 & S2_data_edditting != true)
        {
            document.getElementById(IOField).value = data;
        }
        else if(S2_data_edditting != true)
        {
            document.getElementById(IOField).value = data.toFixed(tofix);
        }
    });
}