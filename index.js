

// Mảng xuất dữ liệu report Excel
var SQL_Excel = [];  // Dữ liệu nhập kho

// /////////////////////////++THIẾT LẬP KẾT NỐI WEB++/////////////////////////
var express = require("express");
var app = express();
app.use(express.static("public"));
app.set("view engine", "ejs");
app.set("views", "./views");
var server = require("http").Server(app);
var io = require("socket.io")(server);
server.listen(8000);
// Home calling
app.get("/", function(req, res){
    res.render("home")
});
//

// KHỞI TẠO KẾT NỐI PLC
var nodes7 = require('nodes7');
var conn_plc = new nodes7; //PLC1
// Tạo địa chỉ kết nối (slot = 2 nếu là 300/400, slot = 1 nếu là 1200/1500)
conn_plc.initiateConnection({port: 102, host: '192.168.1.8', rack: 0, slot: 1}, PLC_connected);


// Bảng tag trong Visual studio code
var tags_list = { 
    B_W_ON_SYSTEM: 'DB100,X0.0',          
    B_W_OFF_SYSTEM: 'DB100,X0.1',        
    B_W_AUTO: 'DB100,X0.2',       
    B_W_MANUAL: 'DB100,X0.3',
    B_W_RA: 'DB100,X0.4',    
    B_W_VAO: 'DB100,X0.5',    
    B_W_LEN: 'DB100,X0.6',    
    B_W_XUONG: 'DB100,X0.7',   
    B_W_QUAY: 'DB100,X1.0',    
    B1_W: 'DB100,X1.1',    
    B2_W: 'DB100,X1.2', 
    B3_W: 'DB100,X1.3', 
    B4_W: 'DB100,X1.4',             
    B5_W: 'DB100,X1.5', 
    B6_W: 'DB100,X1.6', 
    B7_W: 'DB100,X1.7', 
    B8_W: 'DB100,X2.0', 
    B9_W: 'DB100,X2.1', 
    B10_W: 'DB100,X2.2', 
    B11_W: 'DB100,X2.3', 
    B12_W: 'DB100,X2.4', 
    B13_W: 'DB100,X2.5', 
    B14_W: 'DB100,X2.6', 
    B15_W: 'DB100,X2.7', 
    B16_W: 'DB100,X3.0', 
    B17_W: 'DB100,X3.1', 
    B18_W: 'DB100,X3.2', 
    B19_W: 'DB100,X3.3', 
    B20_W: 'DB100,X3.4', 
    B21_W: 'DB100,X3.5', 
    B22_W: 'DB100,X3.6', 
    B23_W: 'DB100,X3.7', 
    B24_W: 'DB100,X4.0', 
    B_IN_W: 'DB100,X4.1', 
    B_OUT_W: 'DB100,X4.2', 
    B_HUY_W: 'DB100,X4.3', 
    B_GUI_W: 'DB100,X4.4', 
    B_LAY_W: 'DB100,X4.5', 
    B_CHUAN_TANG_W: 'DB100,X4.6', 
    B_SET_HOME_W: 'DB100,X4.7', 
    B_PAUSE_RESUME_W: 'DB100,X5.0',
    Q1_W: 'DB100,X5.1', 
    Q2_W: 'DB100,X5.2', 
    Q3_W: 'DB100,X5.3', 
    Q4_W: 'DB100,X5.4', 
    Q5_W: 'DB100,X5.5', 
    Q6_W: 'DB100,X5.6', 
    Q7_W: 'DB100,X5.7', 
    Q8_W: 'DB100,X6.0', 
    Q9_W: 'DB100,X6.1', 
    Q10_W: 'DB100,X6.2', 
    Q11_W: 'DB100,X6.3', 
    Q12_W: 'DB100,X6.4', 
    Q13_W: 'DB100,X6.5', 
    Q14_W: 'DB100,X6.6', 
    Q15_W: 'DB100,X6.7', 
    Q16_W: 'DB100,X7.0', 
    Q17_W: 'DB100,X7.1', 
    Q18_W: 'DB100,X7.2',
    Q19_W: 'DB100,X7.3',
    Q20_W: 'DB100,X7.4',
    Q21_W: 'DB100,X7.5',
    Q22_W: 'DB100,X7.6',
    Q23_W: 'DB100,X7.7',
    Q24_W: 'DB100,X8.0',
    Q_IN__W: 'DB100,X8.1',
    Q_OUT_W: 'DB100,X8.2',
    B_RESET_W: 'DB100,X8.3',
    STT_SYSTEM_W: 'DB100,X8.4',
    BIT_AUTO_W: 'DB100,X8.5',
    BIT_MANUAL_W: 'DB100,X8.6',
    BIT_GUI_W: 'DB100,X8.7',
    BIT_LAY_W: 'DB100,X9.0',
    B_RUN_W: 'DB100,X9.1',
    GOC_QUAY_W: 'DB100,REAL10',
    B_POWER_STEP_W: 'DB100,X14.0',
    B_TT_W_1: 'DB100,X14.1',
    B_TT_W_2: 'DB100,X14.2',
    B_TT_W_3: 'DB100,X14.3',
    B_TT_W_4: 'DB100,X14.4',
    B_TT_W_5: 'DB100,X14.5',
    B_TT_W_6: 'DB100,X14.6',
    B_TT_W_7: 'DB100,X14.7',
    B_TT_W_8: 'DB100,X15.0',
    B_TT_W_9: 'DB100,X15.1',
    B_TT_W_10: 'DB100,X15.2',
    B_TT_W_11: 'DB100,X15.3',
    B_TT_W_12: 'DB100,X15.4',
    B_TT_W_13: 'DB100,X15.5',
    B_TT_W_14: 'DB100,X15.6',
    B_TT_W_15: 'DB100,X15.7',
    B_TT_W_16: 'DB100,X16.0',
    B_TT_W_17: 'DB100,X16.1',
    B_TT_W_18: 'DB100,X16.2',
    B_TT_W_19: 'DB100,X16.3',
    B_TT_W_20: 'DB100,X16.4',
    B_TT_W_21: 'DB100,X16.5',
    B_TT_W_22: 'DB100,X16.6',
    B_TT_W_23: 'DB100,X16.7',
    B_TT_W_24: 'DB100,X17.0',
    SO_CO_XE_W: 'DB100,INT18',
    SO_KO_XE_W: 'DB100,INT20', 
    VT_CHECK_AUTO_W: 'DB100,S22.256',
    VI_TRI_TT_XE_W: 'DB100,INT278', 
    THONGTIN_TRUY_TT0_W: 'DB100,S280.256', 
    THONGTIN_TRUY_TT1_W: 'DB100,S536.256', 
    THONGTIN_TRUY_TT2_W: 'DB100,S792.256', 
    THONGTIN_TRUY_TT3_W: 'DB100,S1048.256', 
    HMI_TT_VAO_RA_TT_W: 'DB100,S1304.256', 
    TT_VAO_RA_TT0_W: 'DB100,S1560.256', 
    TT_VAO_RA_TT1_W: 'DB100,S1816.256', 
    HMI_TT_VAO_RA_C4_W: 'DB100,S2072.256', 
    TT_VAO_RA_TT2_W: 'DB100,DTL2328', 
    HMI_TT_VAO_RA_C6_W: 'DB100,S2340.256', 
    TT_VAO_RA_TT3_W: 'DB100,WORD2596',
    
};


// GỬI DỮ LIỆu TAG CHO PLC
function PLC_connected(err) {
    if (typeof(err) !== "undefined") {
        console.log(err); // Hiển thị lỗi nếu không kết nối đƯỢc với PLC
    }
    conn_plc.setTranslationCB(function(tag) {return tags_list[tag];});  // Đưa giá trị đọc lên từ PLC và mảng
    conn_plc.addItems([      
      'B_W_ON_SYSTEM',      
      'B_W_OFF_SYSTEM',      
      'B_W_AUTO',    
      'B_W_MANUAL',   
      'B_W_RA',     
      'B_W_VAO', 
      'B_W_LEN', 
      'B_W_XUONG', 
      'B_W_QUAY', 
      'B1_W', 
      'B2_W',
      'B3_W',  
      'B4_W', 
      'B5_W', 
      'B6_W', 
      'B7_W', 
      'B8_W', 
      'B9_W', 
      'B10_W', 
      'B11_W', 
      'B12_W', 
      'B13_W', 
      'B14_W', 
      'B15_W', 
      'B16_W', 
      'B17_W', 
      'B18_W', 
      'B19_W', 
      'B20_W', 
      'B21_W', 
      'B22_W', 
      'B23_W', 
      'B24_W', 
      'B_IN_W', 
      'B_OUT_W', 
      'B_HUY_W', 
      'B_GUI_W', 
      'B_LAY_W', 
      'B_CHUAN_TANG_W', 
      'B_SET_HOME_W', 
      'B_PAUSE_RESUME_W',
      'Q1_W',
      'Q2_W',      
      'Q3_W',
      'Q4_W',
      'Q5_W',
      'Q6_W',
      'Q7_W',
      'Q8_W',
      'Q9_W',
      'Q10_W',
      'Q11_W',
      'Q12_W',
      'Q13_W',
      'Q14_W',
      'Q15_W',
      'Q16_W',
      'Q17_W',
      'Q18_W',
      'Q19_W',
      'Q20_W',
      'Q21_W',
      'Q22_W',
      'Q23_W',
      'Q24_W',
      'Q_IN__W',
      'Q_OUT_W',
      'B_RESET_W',
      'STT_SYSTEM_W',
      'BIT_AUTO_W',
      'BIT_MANUAL_W',
      'BIT_GUI_W',
      'BIT_LAY_W',
      'B_RUN_W',
      'GOC_QUAY_W',
      'B_POWER_STEP_W',
      'B_TT_W_1',
      'B_TT_W_2',
      'B_TT_W_3',
      'B_TT_W_4',
      'B_TT_W_5',
      'B_TT_W_6',
      'B_TT_W_7',
      'B_TT_W_8',
      'B_TT_W_9',
      'B_TT_W_10',
      'B_TT_W_11',
      'B_TT_W_12',
      'B_TT_W_13',
      'B_TT_W_14',
      'B_TT_W_15',
      'B_TT_W_16',
      'B_TT_W_17',
      'B_TT_W_18',
      'B_TT_W_19',
      'B_TT_W_20',
      'B_TT_W_21',
      'B_TT_W_22',
      'B_TT_W_23',
      'B_TT_W_24',
      'SO_CO_XE_W',
      'SO_KO_XE_W',
      'VT_CHECK_AUTO_W',
      'VI_TRI_TT_XE_W',
      'THONGTIN_TRUY_TT0_W',
      'THONGTIN_TRUY_TT1_W',
      'THONGTIN_TRUY_TT2_W',
      'THONGTIN_TRUY_TT3_W',
      'HMI_TT_VAO_RA_TT_W',
      'TT_VAO_RA_TT0_W',
      'TT_VAO_RA_TT1_W',
      'HMI_TT_VAO_RA_C4_W',
      'TT_VAO_RA_TT2_W',
      'HMI_TT_VAO_RA_C6_W',
      'TT_VAO_RA_TT3_W'
    ]);
}


// Đọc dữ liệu từ PLC và đưa vào array tags
var arr_tag_value = []; // Tạo một mảng lưu giá trị tag đọc về
function valuesReady(anythingBad, values) {
    if (anythingBad) { console.log("Lỗi khi đọc dữ liệu tag"); } // Cảnh báo lỗi
    var lodash = require('lodash'); // Chuyển variable sang array
    arr_tag_value = lodash.map(values, (item) => item);
    console.log(values); // Hiển thị giá trị để kiểm tra
}

// Hàm chức năng scan giá trị
function fn_read_data_scan(){
    conn_plc.readAllItems(valuesReady);
    
}
// Time cập nhật mỗi 1s
setInterval(

    () => fn_read_data_scan(),
    1000 // 1s = 1000ms
);

// HÀM GHI DỮ LIỆU XUỐNG PLC
function valuesWritten(anythingBad) {
    if (anythingBad) { console.log("WRONG VALUES!"); }
    console.log("Done writing.");
}

// Nhận các bức điện được gửi từ trình duyệt
io.on("connection", function(socket){
    socket.on("Client-send-start", function(data){conn_plc.writeItems('B_W_ON_SYSTEM', data, valuesWritten);});
    socket.on("Client-send-stop", function(data){conn_plc.writeItems('B_W_OFF_SYSTEM', data, valuesWritten);});
    socket.on("Client-send-pr", function(data){conn_plc.writeItems('B_PAUSE_RESUME_W', data, valuesWritten);});
    socket.on("Client-send-auto", function(data){conn_plc.writeItems('B_W_AUTO', data, valuesWritten);});
    socket.on("Client-send-manu", function(data){conn_plc.writeItems('B_W_MANUAL', data, valuesWritten);});
    socket.on("Client-send-park", function(data){conn_plc.writeItems('B_GUI_W', data, valuesWritten);});
    socket.on("Client-send-retrieve", function(data){conn_plc.writeItems('B_LAY_W', data, valuesWritten);});
    socket.on("Client-send-in", function(data){conn_plc.writeItems('B_IN_W', data, valuesWritten);});
    socket.on("Client-send-out", function(data){conn_plc.writeItems('B_OUT_W', data, valuesWritten);});
    socket.on("Client-send-1", function(data){conn_plc.writeItems('B1_W', data, valuesWritten);});
    socket.on("Client-send-2", function(data){conn_plc.writeItems('B2_W', data, valuesWritten);});
    socket.on("Client-send-3", function(data){conn_plc.writeItems('B3_W', data, valuesWritten);});
    socket.on("Client-send-4", function(data){conn_plc.writeItems('B4_W', data, valuesWritten);});
    socket.on("Client-send-5", function(data){conn_plc.writeItems('B5_W', data, valuesWritten);});
    socket.on("Client-send-6", function(data){conn_plc.writeItems('B6_W', data, valuesWritten);});
    socket.on("Client-send-7", function(data){conn_plc.writeItems('B7_W', data, valuesWritten);});
    socket.on("Client-send-8", function(data){conn_plc.writeItems('B8_W', data, valuesWritten);});
    socket.on("Client-send-9", function(data){conn_plc.writeItems('B9_W', data, valuesWritten);});
    socket.on("Client-send-10", function(data){conn_plc.writeItems('B10_W', data, valuesWritten);});
    socket.on("Client-send-11", function(data){conn_plc.writeItems('B11_W', data, valuesWritten);});
    socket.on("Client-send-12", function(data){conn_plc.writeItems('B12_W', data, valuesWritten);});
    socket.on("Client-send-13", function(data){conn_plc.writeItems('B13_W', data, valuesWritten);});
    socket.on("Client-send-14", function(data){conn_plc.writeItems('B14_W', data, valuesWritten);});
    socket.on("Client-send-15", function(data){conn_plc.writeItems('B15_W', data, valuesWritten);});
    socket.on("Client-send-16", function(data){conn_plc.writeItems('B16_W', data, valuesWritten);});
    socket.on("Client-send-17", function(data){conn_plc.writeItems('B17_W', data, valuesWritten);});
    socket.on("Client-send-18", function(data){conn_plc.writeItems('B18_W', data, valuesWritten);});
    socket.on("Client-send-19", function(data){conn_plc.writeItems('B19_W', data, valuesWritten);});
    socket.on("Client-send-20", function(data){conn_plc.writeItems('B20_W', data, valuesWritten);});
    socket.on("Client-send-21", function(data){conn_plc.writeItems('B21_W', data, valuesWritten);});
    socket.on("Client-send-22", function(data){conn_plc.writeItems('B22_W', data, valuesWritten);});
    socket.on("Client-send-23", function(data){conn_plc.writeItems('B23_W', data, valuesWritten);});
    socket.on("Client-send-24", function(data){conn_plc.writeItems('B24_W', data, valuesWritten);});
    socket.on("Client-send-power", function(data){conn_plc.writeItems('B_POWER_STEP_W', data, valuesWritten);});
    socket.on("Client-send-shome", function(data){conn_plc.writeItems('B_SET_HOME_W', data, valuesWritten);});
    socket.on("Client-send-chuan", function(data){conn_plc.writeItems('B_CHUAN_TANG_W', data, valuesWritten);});
    socket.on("Client-send-len", function(data){conn_plc.writeItems('B_W_LEN', data, valuesWritten);});
    socket.on("Client-send-xuong", function(data){conn_plc.writeItems('B_W_XUONG', data, valuesWritten);});
    socket.on("Client-send-quay", function(data){conn_plc.writeItems('B_W_QUAY', data, valuesWritten);});
    socket.on("Client-send-ra", function(data){conn_plc.writeItems('B_W_RA', data, valuesWritten);});
    socket.on("Client-send-vo", function(data){conn_plc.writeItems('B_W_VAO', data, valuesWritten);});
    socket.on("Client-send-huy", function(data){conn_plc.writeItems('B_HUY_W', data, valuesWritten);});
    socket.on("Client-send-T-1", function(data){conn_plc.writeItems('B_TT_W_1', data, valuesWritten);});
    socket.on("Client-send-T-2", function(data){conn_plc.writeItems('B_TT_W_2', data, valuesWritten);});
    socket.on("Client-send-T-3", function(data){conn_plc.writeItems('B_TT_W_3', data, valuesWritten);});
    socket.on("Client-send-T-4", function(data){conn_plc.writeItems('B_TT_W_4', data, valuesWritten);});
    socket.on("Client-send-T-5", function(data){conn_plc.writeItems('B_TT_W_5', data, valuesWritten);});
    socket.on("Client-send-T-6", function(data){conn_plc.writeItems('B_TT_W_6', data, valuesWritten);});
    socket.on("Client-send-T-7", function(data){conn_plc.writeItems('B_TT_W_7', data, valuesWritten);});
    socket.on("Client-send-T-8", function(data){conn_plc.writeItems('B_TT_W_8', data, valuesWritten);});
    socket.on("Client-send-T-9", function(data){conn_plc.writeItems('B_TT_W_9', data, valuesWritten);});
    socket.on("Client-send-T-10", function(data){conn_plc.writeItems('B_TT_W_10', data, valuesWritten);});
    socket.on("Client-send-T-11", function(data){conn_plc.writeItems('B_TT_W_11', data, valuesWritten);});
    socket.on("Client-send-T-12", function(data){conn_plc.writeItems('B_TT_W_12', data, valuesWritten);});
    socket.on("Client-send-T-13", function(data){conn_plc.writeItems('B_TT_W_13', data, valuesWritten);});
    socket.on("Client-send-T-14", function(data){conn_plc.writeItems('B_TT_W_14', data, valuesWritten);});
    socket.on("Client-send-T-15", function(data){conn_plc.writeItems('B_TT_W_15', data, valuesWritten);});
    socket.on("Client-send-T-16", function(data){conn_plc.writeItems('B_TT_W_16', data, valuesWritten);});
    socket.on("Client-send-T-17", function(data){conn_plc.writeItems('B_TT_W_17', data, valuesWritten);});
    socket.on("Client-send-T-18", function(data){conn_plc.writeItems('B_TT_W_18', data, valuesWritten);});
    socket.on("Client-send-T-19", function(data){conn_plc.writeItems('B_TT_W_19', data, valuesWritten);});
    socket.on("Client-send-T-20", function(data){conn_plc.writeItems('B_TT_W_20', data, valuesWritten);});
    socket.on("Client-send-T-21", function(data){conn_plc.writeItems('B_TT_W_21', data, valuesWritten);});
    socket.on("Client-send-T-22", function(data){conn_plc.writeItems('B_TT_W_22', data, valuesWritten);});
    socket.on("Client-send-T-23", function(data){conn_plc.writeItems('B_TT_W_23', data, valuesWritten);});
    socket.on("Client-send-T-24", function(data){conn_plc.writeItems('B_TT_W_24', data, valuesWritten);});
    socket.on("Client-send-run", function(data){conn_plc.writeItems('B_RUN_W', data, valuesWritten);});
    socket.on("Client-send-reset", function(data){conn_plc.writeItems('B_RESET_W', data, valuesWritten);});
});


// ///////////LẬP BẢNG TAG ĐỂ GỬI QUA CLIENT (TRÌNH DUYỆT)///////////
function fn_tag(){
    io.sockets.emit("B_W_ON_SYSTEM", arr_tag_value[0]);
    io.sockets.emit("B_W_OFF_SYSTEM", arr_tag_value[1]);
    io.sockets.emit("B_W_AUTO", arr_tag_value[2]);
    io.sockets.emit("B_W_MANUAL", arr_tag_value[3]);
    io.sockets.emit("B_W_RA", arr_tag_value[4]);
    io.sockets.emit("B_W_VAO", arr_tag_value[5]);
    io.sockets.emit("B_W_LEN", arr_tag_value[6]);
    io.sockets.emit("B_W_XUONG", arr_tag_value[7]);
    io.sockets.emit("B_W_QUAY", arr_tag_value[8]);
    io.sockets.emit("B1_W", arr_tag_value[9]);
    io.sockets.emit("B2_W", arr_tag_value[10]);
    io.sockets.emit("B3_W", arr_tag_value[11]);
    io.sockets.emit("B4_W", arr_tag_value[12]);
    io.sockets.emit("B5_W", arr_tag_value[13]);
    io.sockets.emit("B6_W", arr_tag_value[14]);
    io.sockets.emit("B7_W", arr_tag_value[15]);
    io.sockets.emit("B8_W", arr_tag_value[16]);
    io.sockets.emit("B9_W", arr_tag_value[17]);
    io.sockets.emit("B10_W", arr_tag_value[18]);
    io.sockets.emit("B11_W", arr_tag_value[19]);
    io.sockets.emit("B12_W", arr_tag_value[20]);
    io.sockets.emit("B13_W", arr_tag_value[21]);
    io.sockets.emit("B14_W", arr_tag_value[22]);
    io.sockets.emit("B15_W", arr_tag_value[23]);
    io.sockets.emit("B16_W", arr_tag_value[24]);
    io.sockets.emit("B17_W", arr_tag_value[25]);
    io.sockets.emit("B18_W", arr_tag_value[26]);
    io.sockets.emit("B19_W", arr_tag_value[27]);
    io.sockets.emit("B20_W", arr_tag_value[28]);
    io.sockets.emit("B21_W", arr_tag_value[29]);
    io.sockets.emit("B22_W", arr_tag_value[30]);
    io.sockets.emit("B23_W", arr_tag_value[31]);
    io.sockets.emit("B24_W", arr_tag_value[32]);
    io.sockets.emit("B_IN_W", arr_tag_value[33]);
    io.sockets.emit("B_OUT_W", arr_tag_value[34]);
    io.sockets.emit("B_HUY_W", arr_tag_value[35]);
    io.sockets.emit("B_GUI_W", arr_tag_value[36]);
    io.sockets.emit("B_LAY_W", arr_tag_value[37]);
    io.sockets.emit("B_CHUAN_TANG_W", arr_tag_value[38]);
    io.sockets.emit("B_SET_HOME_W", arr_tag_value[39]);
    io.sockets.emit("B_PAUSE_RESUME_W", arr_tag_value[40]);
    io.sockets.emit("Q1_W", arr_tag_value[41]);
    io.sockets.emit("Q2_W", arr_tag_value[42]);
    io.sockets.emit("Q3_W", arr_tag_value[43]);
    io.sockets.emit("Q4_W", arr_tag_value[44]);
    io.sockets.emit("Q5_W", arr_tag_value[45]);
    io.sockets.emit("Q6_W", arr_tag_value[46]);
    io.sockets.emit("Q7_W", arr_tag_value[47]);
    io.sockets.emit("Q8_W", arr_tag_value[48]);
    io.sockets.emit("Q9_W", arr_tag_value[49]);
    io.sockets.emit("Q10_W", arr_tag_value[50]);
    io.sockets.emit("Q11_W", arr_tag_value[51]);
    io.sockets.emit("Q12_W", arr_tag_value[52]);
    io.sockets.emit("Q13_W", arr_tag_value[53]);
    io.sockets.emit("Q14_W", arr_tag_value[54]);
    io.sockets.emit("Q15_W", arr_tag_value[55]);
    io.sockets.emit("Q16_W", arr_tag_value[56]);
    io.sockets.emit("Q17_W", arr_tag_value[57]);
    io.sockets.emit("Q18_W", arr_tag_value[58]);
    io.sockets.emit("Q19_W", arr_tag_value[59]);
    io.sockets.emit("Q20_W", arr_tag_value[60]);
    io.sockets.emit("Q21_W", arr_tag_value[61]);
    io.sockets.emit("Q22_W", arr_tag_value[62]);
    io.sockets.emit("Q23_W", arr_tag_value[63]);
    io.sockets.emit("Q24_W", arr_tag_value[64]);
    io.sockets.emit("Q_IN__W", arr_tag_value[65]);
    io.sockets.emit("Q_OUT_W", arr_tag_value[66]);
    io.sockets.emit("B_RESET_W", arr_tag_value[67]);
    io.sockets.emit("STT_SYSTEM_W", arr_tag_value[68]);
    io.sockets.emit("BIT_AUTO_W", arr_tag_value[69]);
    io.sockets.emit("BIT_MANUAL_W", arr_tag_value[70]);
    io.sockets.emit("BIT_GUI_W", arr_tag_value[71]);
    io.sockets.emit("BIT_LAY_W", arr_tag_value[72]);
    io.sockets.emit("B_RUN_W", arr_tag_value[73]);
    io.sockets.emit("GOC_QUAY_W", arr_tag_value[74]);
    io.sockets.emit("B_POWER_STEP_W", arr_tag_value[75]);
    io.sockets.emit("B_TT_W_1", arr_tag_value[76]);
    io.sockets.emit("B_TT_W_2", arr_tag_value[77]);
    io.sockets.emit("B_TT_W_3", arr_tag_value[78]);
    io.sockets.emit("B_TT_W_4", arr_tag_value[79]);
    io.sockets.emit("B_TT_W_5", arr_tag_value[80]);
    io.sockets.emit("B_TT_W_6", arr_tag_value[81]);
    io.sockets.emit("B_TT_W_7", arr_tag_value[82]);
    io.sockets.emit("B_TT_W_8", arr_tag_value[83]);
    io.sockets.emit("B_TT_W_9", arr_tag_value[84]);
    io.sockets.emit("B_TT_W_10", arr_tag_value[85]);
    io.sockets.emit("B_TT_W_11", arr_tag_value[86]);
    io.sockets.emit("B_TT_W_12", arr_tag_value[87]);
    io.sockets.emit("B_TT_W_13", arr_tag_value[88]);
    io.sockets.emit("B_TT_W_14", arr_tag_value[89]);
    io.sockets.emit("B_TT_W_15", arr_tag_value[90]);
    io.sockets.emit("B_TT_W_16", arr_tag_value[91]);
    io.sockets.emit("B_TT_W_17", arr_tag_value[92]);
    io.sockets.emit("B_TT_W_18", arr_tag_value[93]);
    io.sockets.emit("B_TT_W_19", arr_tag_value[94]);
    io.sockets.emit("B_TT_W_20", arr_tag_value[95]);
    io.sockets.emit("B_TT_W_21", arr_tag_value[96]);
    io.sockets.emit("B_TT_W_22", arr_tag_value[97]);
    io.sockets.emit("B_TT_W_23", arr_tag_value[98]);
    io.sockets.emit("B_TT_W_24", arr_tag_value[99]);
    io.sockets.emit("SO_CO_XE_W", arr_tag_value[100]);
    io.sockets.emit("SO_KO_XE_W", arr_tag_value[101]);
    io.sockets.emit("VT_CHECK_AUTO_W", arr_tag_value[102]);
    io.sockets.emit("VI_TRI_TT_XE_W", arr_tag_value[103]);
    io.sockets.emit("THONGTIN_TRUY_TT0_W", arr_tag_value[104]);
    io.sockets.emit("THONGTIN_TRUY_TT1_W", arr_tag_value[105]);
    io.sockets.emit("THONGTIN_TRUY_TT2_W", arr_tag_value[106]);
    io.sockets.emit("THONGTIN_TRUY_TT3_W", arr_tag_value[107]);
    io.sockets.emit("HMI_TT_VAO_RA_TT_W", arr_tag_value[108]);
    io.sockets.emit("TT_VAO_RA_TT0_W", arr_tag_value[109]);
    io.sockets.emit("TT_VAO_RA_TT1_W", arr_tag_value[110]);
    io.sockets.emit("HMI_TT_VAO_RA_C4_W", arr_tag_value[111]);
    io.sockets.emit("TT_VAO_RA_TT2_W", arr_tag_value[112]);
    io.sockets.emit("HMI_TT_VAO_RA_C6_W", arr_tag_value[113]);
    io.sockets.emit("TT_VAO_RA_TT3_W", arr_tag_value[114]);
}
// ///////////GỬI DỮ LIỆU BẢNG TAG ĐẾN CLIENT (TRÌNH DUYỆT)///////////
io.on("connection", function(socket){
    socket.on("Client-send-data", function(data){
    fn_tag();
    });
    fn_SQLSearch(); // Hàm tìm kiếm SQL
    fn_SQLSearch_ByTime();
    fn_SQLSearch_ByTime_In();
    fn_SQLSearch_ByTime_Out();
    fn_SQLSearch_ByNum_Car();
    fn_SQLSearch_ByColor_Car();
    fn_SQLSearch_ByPosition_Car();
    fn_SQLSearch_KHTT();
    fn_Require_ExcelExport(); // Nhận yêu cầu xuất Excel
});



// Khởi tạo SQL
var mysql = require('mysql');
var sqlcon = mysql.createConnection({
  host: "localhost",
  user: "root",
  password: "123456",
  database: "SQL_PLC",
  dateStrings:true // Hiển thị không có T và Z
});


// Đọc dữ liệu từ SQL
function fn_SQLSearch(){
    io.on("connection", function(socket){
        socket.on("msg_SQL_Show", function(data)
        {
            var sqltable_Name = "parking_auto";
            var queryy1 = "SELECT * FROM " + sqltable_Name + ";"
            sqlcon.query(queryy1, function(err, results, fields) {
                if (err) {
                    console.log(err);
                } else {
                    const objectifyRawPacket = row => ({...row});
                    const convertedResponse = results.map(objectifyRawPacket);
                    SQL_Excel = convertedResponse; // Xuất báo cáo Excel
                    socket.emit('SQL_Show', convertedResponse);
                    console.log(convertedResponse);
                }
            });
        });
    });
    }

// Đọc dữ liệu từ SQL
function fn_SQLSearch_KHTT(){
    io.on("connection", function(socket){
        socket.on("msg_SQL_KHTT", function(data)
        {
            var sqltable_Name = "parking_auto";
            var dt_col_Name = "Member";  // Tên cột thời gian
            var queryy1 = "SELECT * FROM " + sqltable_Name + " WHERE "+ dt_col_Name +  "='Khách hàng thân thiết';";
            sqlcon.query(queryy1, function(err, results, fields) {
                if (err) {
                    console.log(err);
                } else {
                    const objectifyRawPacket = row => ({...row});
                    const convertedResponse = results.map(objectifyRawPacket);
                    SQL_Excel = convertedResponse; // Xuất báo cáo Excel
                    socket.emit('SQL_KHTT', convertedResponse);
                    console.log(convertedResponse);
                }
            });
        });
    });
    }

// Đọc dữ liệu SQL theo thời gian
function fn_SQLSearch_ByTime(){
    io.on("connection", function(socket){
        socket.on("msg_SQL_ByTime", function(data)
        {
            var tzoffset = (new Date()).getTimezoneOffset() * 60000; //offset time Việt Nam (GMT7+)

            var timeS = new Date(data[0]); // Thời gian bắt đầu
            var timeE = new Date(data[1]);

            var timess = new Date(timeS - tzoffset);
            var timerr = new Date(timeE - tzoffset)

            const [date1, time1] = timess.split('T');
            const [year1, month1, day1] = date1.split('-');
            const formattedDatetime1 = `${time1} ${day1}-${month1}-${year1}`;

            const [date2, time2] = timerr.split('T');
            const [year2, month2, day2] = date2.split('-');
            const formattedDatetime2 = `${time2} ${day2}-${month2}-${year2}`;

            // Quy đổi thời gian ra định dạng cua MySQL
            var timeS1 = "'" + formattedDatetime1 + "'";
            var timeE1 = "'" + formattedDatetime2 + "'";
            var timeR = timeS1 + "AND" + timeE1; // Khoảng thời gian tìm kiếm (Time Range)

            var sqltable_Name = "parking_auto"; // Tên bảng
            var dt_col_Name = "Time_In";  // Tên cột thời gian

            var Query1 = "SELECT * FROM " + sqltable_Name + " WHERE "+ dt_col_Name + " BETWEEN ";
            var Query = Query1 + timeR + ";";

            sqlcon.query(Query, function(err, results, fields) {
                if (err) {
                    console.log(err);
                } else {
                    const objectifyRawPacket = row => ({...row});
                    const convertedResponse = results.map(objectifyRawPacket);
                    SQL_Excel = convertedResponse; // Xuất báo cáo Excel
                    socket.emit('SQL_ByTime', convertedResponse);
                }
            });
        });
    });
}

function formatDateTime(datetime) {
    const date = new Date(datetime);
    const hours = date.getHours().toString().padStart(2, '0');
    const minutes = date.getMinutes().toString().padStart(2, '0');
    const day = date.getDate().toString().padStart(2, '0');
    const month = (date.getMonth() + 1).toString().padStart(2, '0'); // Months are 0-based in JavaScript
    const year = date.getFullYear();
    return `${hours}:${minutes} ${day}-${month}-${year}`;
}

// Đọc dữ liệu SQL theo thời gian
function fn_SQLSearch_ByTime_In1(){
    io.on("connection", function(socket){
        socket.on("msg_SQL_ByTime_In", function(data)
        {
            var tzoffset = (new Date()).getTimezoneOffset() * 60000; //offset time Việt Nam (GMT7+)
            // Lấy thời gian tìm kiếm từ date time piker
            var timeS = new Date(data[0]); // Thời gian bắt đầu
            var timeE = new Date(data[1]); // Thời gian kết thúc
            // Quy đổi thời gian ra định dạng cua MySQL
            var timeS1 = "'" + (new Date(timeS - tzoffset)).toISOString().slice(0, -1).replace("T"," ")	+ "'";
            var timeE1 = "'" + (new Date(timeE - tzoffset)).toISOString().slice(0, -1).replace("T"," ") + "'";
            var timeR = timeS1 + "AND" + timeE1; // Khoảng thời gian tìm kiếm (Time Range)

            var sqltable_Name = "parking_auto"; // Tên bảng
            var dt_col_Name = "Time_In";  // Tên cột thời gian

            var Query1 = "SELECT * FROM " + sqltable_Name + " WHERE "+ dt_col_Name + " BETWEEN ";
            var Query = Query1 + timeR + ";";

            sqlcon.query(Query, function(err, results, fields) {
                if (err) {
                    console.log(err);
                } else {
                    const objectifyRawPacket = row => ({...row});
                    const convertedResponse = results.map(objectifyRawPacket);
                    SQL_Excel = convertedResponse; // Xuất báo cáo Excel
                    socket.emit('SQL_ByTime_In', convertedResponse);
                }
            });
        });
    });
}


// Đọc dữ liệu SQL theo thời gian
function fn_SQLSearch_ByTime_In(){
    io.on("connection", function(socket){
        socket.on("msg_SQL_ByTime_In", function(data)
        {
            var tzoffset = (new Date()).getTimezoneOffset() * 60000; //offset time Việt Nam (GMT7+)
            // Lấy thời gian tìm kiếm từ date time piker
            var timeS = new Date(data[0]); // Thời gian bắt đầu
            var timeE = new Date(data[1]); // Thời gian kết thúc
            // Quy đổi thời gian ra định dạng cua MySQL
            var timeS1 = "'" + (new Date(timeS - tzoffset)).toISOString().slice(0, -1).replace("T"," ")	+ "'";
            var timeE1 = "'" + (new Date(timeE - tzoffset)).toISOString().slice(0, -1).replace("T"," ") + "'";
            var timeR = timeS1 + "AND" + timeE1; // Khoảng thời gian tìm kiếm (Time Range)

            var sqltable_Name = "parking_auto"; // Tên bảng
            var dt_col_Name = "Time_In";  // Tên cột thời gian

            var Query1 = "SELECT * FROM " + sqltable_Name + " WHERE "+ dt_col_Name + " BETWEEN ";
            var Query = Query1 + timeR + ";";

            sqlcon.query(Query, function(err, results, fields) {
                if (err) {
                    console.log(err);
                } else {
                    const objectifyRawPacket = row => ({...row});
                    const convertedResponse = results.map(objectifyRawPacket);
                    SQL_Excel = convertedResponse; // Xuất báo cáo Excel
                    socket.emit('SQL_ByTime_In', convertedResponse);
                }
            });
        });
    });
}

// Đọc dữ liệu SQL theo thời gian
function fn_SQLSearch_ByTime_Out(){
    io.on("connection", function(socket){
        socket.on("msg_SQL_ByTime_Out", function(data)
        {
            var tzoffset = (new Date()).getTimezoneOffset() * 60000; //offset time Việt Nam (GMT7+)
            // Lấy thời gian tìm kiếm từ date time piker
            var timeS = new Date(data[0]); // Thời gian bắt đầu
            var timeE = new Date(data[1]); // Thời gian kết thúc
            // Quy đổi thời gian ra định dạng cua MySQL
            var timeS1 = "'" + (new Date(timeS - tzoffset)).toISOString().slice(0, -1).replace("T"," ")	+ "'";
            var timeE1 = "'" + (new Date(timeE - tzoffset)).toISOString().slice(0, -1).replace("T"," ") + "'";
            var timeR = timeS1 + "AND" + timeE1; // Khoảng thời gian tìm kiếm (Time Range)

            var sqltable_Name = "parking_auto"; // Tên bảng
            var dt_col_Name = "Time_Out";  // Tên cột thời gian

            var Query1 = "SELECT * FROM " + sqltable_Name + " WHERE "+ dt_col_Name + " BETWEEN ";
            var Query = Query1 + timeR + ";";

            sqlcon.query(Query, function(err, results, fields) {
                if (err) {
                    console.log(err);
                } else {
                    const objectifyRawPacket = row => ({...row});
                    const convertedResponse = results.map(objectifyRawPacket);
                    SQL_Excel = convertedResponse; // Xuất báo cáo Excel
                    socket.emit('SQL_ByTime_Out', convertedResponse);
                }
            });
        });
    });
}

// Đọc dữ liệu SQL theo thời gian
function fn_SQLSearch_ByNum_Car(){
    io.on("connection", function(socket){
        socket.on("msg_SQL_ByNum_Car", function(data)
        {
            var Numcar = data;
            var sqltable_Name = "parking_auto"; // Tên bảng
            var dt_col_Name = "Num";  // Tên cột thời gian

            var Query = "SELECT * FROM " + sqltable_Name + " WHERE "+ dt_col_Name + "='" + Numcar + "';";

            sqlcon.query(Query, function(err, results, fields) {
                if (err) {
                    console.log(err);
                } else {
                    const objectifyRawPacket = row => ({...row});
                    const convertedResponse = results.map(objectifyRawPacket);
                    SQL_Excel = convertedResponse; // Xuất báo cáo Excel
                    socket.emit('SQL_ByNum_Car', convertedResponse);
                }
            });
        });
    });
}

// Đọc dữ liệu SQL theo thời gian
function fn_SQLSearch_ByColor_Car(){
    io.on("connection", function(socket){
        socket.on("msg_SQL_ByColor_Car", function(data)
        {
            var Colorcar = data;
            var sqltable_Name = "parking_auto"; // Tên bảng
            var dt_col_Name = "Color";  // Tên cột thời gian

            var Query = "SELECT * FROM " + sqltable_Name + " WHERE "+ dt_col_Name + "='" + Colorcar + "';";

            sqlcon.query(Query, function(err, results, fields) {
                if (err) {
                    console.log(err);
                } else {
                    const objectifyRawPacket = row => ({...row});
                    const convertedResponse = results.map(objectifyRawPacket);
                    SQL_Excel = convertedResponse; // Xuất báo cáo Excel
                    socket.emit('SQL_ByColor_Car', convertedResponse);
                }
            });
        });
    });
}

// Đọc dữ liệu SQL theo thời gian
function fn_SQLSearch_ByPosition_Car(){
    io.on("connection", function(socket){
        socket.on("msg_SQL_ByPosition_Car", function(data)
        {
            var Positioncar = data;
            var sqltable_Name = "parking_auto"; // Tên bảng
            var dt_col_Name = "Position";  // Tên cột thời gian

            var Query = "SELECT * FROM " + sqltable_Name + " WHERE "+ dt_col_Name + "='" + Positioncar + "';";

            sqlcon.query(Query, function(err, results, fields) {
                if (err) {
                    console.log(err);
                } else {
                    const objectifyRawPacket = row => ({...row});
                    const convertedResponse = results.map(objectifyRawPacket);
                    SQL_Excel = convertedResponse; // Xuất báo cáo Excel
                    socket.emit('SQL_ByPosition_Car', convertedResponse);
                }
            });
        });
    });
}

io.on("connection", function(socket)
{
    // Ghi dữ liệu từ IO field
    socket.on("cmd_S2_Edit_Data", function(data){
        conn_plc.writeItems([
                            'GOC_QUAY_W'],
                            [data[0]
                        ], valuesWritten);  
        });
});


// /////////////////////////////// BÁO CÁO EXCEL ///////////////////////////////
const Excel = require('exceljs');
const { CONNREFUSED } = require('dns');
function fn_excelExport(){
    // =====================CÁC THUỘC TÍNH CHUNG=====================
        // Lấy ngày tháng hiện tại
        let date_ob = new Date();
        let date = ("0" + date_ob.getDate()).slice(-2);
        let month = ("0" + (date_ob.getMonth() + 1)).slice(-2);
        let year = date_ob.getFullYear();
        let hours = date_ob.getHours();
        let minutes = date_ob.getMinutes();
        let seconds = date_ob.getSeconds();
        let day = date_ob.getDay();
        var dayName = '';
        if(day == 0){dayName = 'Chủ nhật,'}
        else if(day == 1){dayName = 'Thứ hai,'}
        else if(day == 2){dayName = 'Thứ ba,'}
        else if(day == 3){dayName = 'Thứ tư,'}
        else if(day == 4){dayName = 'Thứ năm,'}
        else if(day == 5){dayName = 'Thứ sáu,'}
        else if(day == 6){dayName = 'Thứ bảy,'}
        else{};
    // Tạo và khai báo Excel
    let workbook = new Excel.Workbook()
    let worksheet =  workbook.addWorksheet('Báo cáo', {
      pageSetup:{paperSize: 9, orientation:'landscape'},
      properties:{tabColor:{argb:'FFC0000'}},
    });
    // Page setup (cài đặt trang)
    worksheet.properties.defaultRowHeight = 20;
    worksheet.pageSetup.margins = {
      left: 0.3, right: 0.25,
      top: 0.75, bottom: 0.75,
      header: 0.3, footer: 0.3
    };
    // =====================THẾT KẾ HEADER=====================
    // Logo công ty
    //const imageId1 = workbook.addImage({
     //   filename: 'public/images/Logo.png',
     //   extension: 'png',
     // });
    //worksheet.addImage(imageId1, 'A1:A3');-->
    // Thông tin công ty
    worksheet.getCell('A1').value = 'Trường Đại Học Sư Phạm Kỹ Thuật TPHCM';
    worksheet.mergeCells('A1:J1');
    worksheet.getCell('A1').style = { font:{bold: true,size: 16},alignment: {horizontal:'center',vertical: 'middle'}} ;
    worksheet.getCell('A2').value = 'Địa chỉ:  Số 1 Võ Văn Ngân, Quận Thủ Đức, TP HCM';
    worksheet.mergeCells('A2:J2');
    worksheet.getCell('A2').style = { alignment: {horizontal:'center',vertical: 'middle'}} ;
    worksheet.getCell('A3').value = 'Hotline: + 0999 999 999';
    worksheet.mergeCells('A3:J3');
    worksheet.getCell('A3').style = { alignment: {horizontal:'center',vertical: 'middle'}} ;
    // Tên báo cáo
    worksheet.getCell('A5').value = 'BÁO CÁO';
    worksheet.mergeCells('A5:J5');
    worksheet.getCell('A5').style = { font:{name: 'Times New Roman', bold: true,size: 16},alignment: {horizontal:'center',vertical: 'middle'}} ;
    // Ngày in biểu
    worksheet.getCell('J6').value = "Ngày in biểu: " + dayName + date + "/" + month + "/" + year + " " + hours + ":" + minutes + ":" + seconds;
    worksheet.getCell('J6').style = { font:{bold: false, italic: true},alignment: {horizontal:'right',vertical: 'bottom',wrapText: false}} ;
    
    // Tên nhãn các cột
    var rowpos = 7;
    var collumName = ["STT","ID", "Num", "Color", "Time In", "Time Out", "Position", "Price", "Member", "Ghi chú"]
    worksheet.spliceRows(rowpos, 1, collumName);
    
    // =====================XUẤT DỮ LIỆU EXCEL SQL=====================
    // Dump all the data into Excel
    var rowIndex = 0;
    SQL_Excel.forEach((e, index) => {
    // row 1 is the header.
    rowIndex =  index + rowpos;
    // worksheet1 collum
    worksheet.columns = [
          {key: 'STT'},
          {key: 'ID'},
          {key: 'Num'},
          {key: 'Color'},
          {key: 'Time_In'},
          {key: 'Time_Out'},
          {key: 'Position'},
          {key: 'Price'},
          {key: 'Member'}
        ]
    worksheet.addRow({
          STT: {
            formula: index + 1
          },
          ...e
        })
    })
    // Lấy tổng số hàng
    const totalNumberOfRows = worksheet.rowCount;
    
    // =====================STYLE CHO CÁC CỘT/HÀNG=====================
    // Style các cột nhãn
    const HeaderStyle = ['A','B', 'C', 'D', 'E','F','G','H','I',,'J']
    HeaderStyle.forEach((v) => {
        worksheet.getCell(`${v}${rowpos}`).style = { font:{bold: true},alignment: {horizontal:'center',vertical: 'middle',wrapText: true}} ;
        worksheet.getCell(`${v}${rowpos}`).border = {
          top: {style:'thin'},
          left: {style:'thin'},
          bottom: {style:'thin'},
          right: {style:'thin'}
        }
    })
    // Cài đặt độ rộng cột
    worksheet.columns.forEach((column, index) => {
        column.width = 15;
    })
    // Set width header
    worksheet.getColumn(1).width = 12;
    worksheet.getColumn(2).width = 12;
    worksheet.getColumn(3).width = 20;
    worksheet.getColumn(5).width = 30;
    worksheet.getColumn(6).width = 30;
    worksheet.getColumn(7).width = 20;
    worksheet.getColumn(8).width = 20;
    worksheet.getColumn(9).width = 35;
    worksheet.getColumn(10).width = 35;
    // ++++++++++++Style cho các hàng dữ liệu++++++++++++
    worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      var datastartrow = rowpos;
      var rowindex = rowNumber + datastartrow;
      const rowlength = datastartrow + SQL_Excel.length
      if(rowindex >= rowlength+1){rowindex = rowlength+1}
      const insideColumns = ['A','B', 'C', 'D', 'E','F','G','H','I','J']
    // Tạo border
      insideColumns.forEach((v) => {
          // Border
        worksheet.getCell(`${v}${rowindex}`).border = {
          top: {style: 'thin'},
          bottom: {style: 'thin'},
          left: {style: 'thin'},
          right: {style: 'thin'}
        },
        // Alignment
        worksheet.getCell(`${v}${rowindex}`).alignment = {horizontal:'center',vertical: 'middle',wrapText: true}
      })
    })
    
    
    // =====================THẾT KẾ FOOTER=====================
    worksheet.getCell(`J${totalNumberOfRows+2}`).value = 'Ngày …………tháng ……………năm 20………';
    worksheet.getCell(`J${totalNumberOfRows+2}`).style = { font:{bold: true, italic: false},alignment: {horizontal:'right',vertical: 'middle',wrapText: false}} ;
    
    worksheet.getCell(`B${totalNumberOfRows+3}`).value = 'Người thực hiện';
    worksheet.getCell(`B${totalNumberOfRows+4}`).value = '(Ký, ghi rõ họ tên)';
    worksheet.getCell(`B${totalNumberOfRows+3}`).style = { font:{bold: true, italic: false},alignment: {horizontal:'center',vertical: 'bottom',wrapText: false}} ;
    worksheet.getCell(`B${totalNumberOfRows+4}`).style = { font:{bold: false, italic: true},alignment: {horizontal:'center',vertical: 'top',wrapText: false}} ;
    
    worksheet.getCell(`J${totalNumberOfRows+3}`).value = 'Người in biểu';
    worksheet.getCell(`J${totalNumberOfRows+4}`).value = '(Ký, ghi rõ họ tên)';
    worksheet.getCell(`J${totalNumberOfRows+3}`).style = { font:{bold: true, italic: false},alignment: {horizontal:'center',vertical: 'bottom',wrapText: false}} ;
    worksheet.getCell(`J${totalNumberOfRows+4}`).style = { font:{bold: false, italic: true},alignment: {horizontal:'center',vertical: 'top',wrapText: false}} ;
    
    
    
    // =====================THỰC HIỆN XUẤT DỮ LIỆU EXCEL=====================
    // Export Link
    var currentTime = year + "_" + month + "_" + date + "_" + hours + "h" + minutes + "m" + seconds + "s";
    var saveasDirect = "Report/Report_" + currentTime + ".xlsx";
    SaveAslink = saveasDirect; // Send to client
    var booknameLink = "public/" + saveasDirect;
    
    var Bookname = "Report_" + currentTime + ".xlsx";
    // Write book name
    workbook.xlsx.writeFile(booknameLink)
    
    // Return
    return [SaveAslink, Bookname]
    
    } // Đóng fn_excelExport

// =====================TRUYỀN NHẬN DỮ LIỆU VỚI TRÌNH DUYỆT=====================
// Hàm chức năng truyền nhận dữ liệu với trình duyệt
function fn_Require_ExcelExport(){
    io.on("connection", function(socket){
        socket.on("msg_Excel_Report", function(data)
        {
            const [SaveAslink1, Bookname] = fn_excelExport();
            var data = [SaveAslink1, Bookname];
            socket.emit('send_Excel_Report', data);
        });
    });
}


