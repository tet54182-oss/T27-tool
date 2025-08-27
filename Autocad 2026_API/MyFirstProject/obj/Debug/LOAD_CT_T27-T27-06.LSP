(defun c:LOAD_CT_AI ( / )
  ;; Tạm thời tắt các thiết lập bảo mật và giao diện dòng lệnh để tải tệp
  (command "_SECURELOAD" "0")
  (command "_CLIPROMPTLINES" "0")
  
  ;; Tải tệp .dll cần thiết
  (command "_netload" "Y:/5.SOFT T27/1. FOR WORK/1. THIET KE DUONG/2.CIVIL 3D/2026/AutoCAD Civil 3D 2026 Win x64/x64/c3d/Civil3D2026.dll")
  
  ;; Khôi phục lại các thiết lập bảo mật và giao diện dòng lệnh
  (command "_SECURELOAD" "1")
  (command "_CLIPROMPTLINES" "2")
  
  ;; Hiển thị thông báo đang tải để người dùng biết quá trình đang diễn ra
  (repeat 250 (princ "\nLoading AECC Land Manager Items...\nLoading AECC Hydrology Manager Items...\nLoading AECC Pipe Network Manager Items...\nLoading AECC Pressure Pipes Manager Items...\nLoading AECC Roadway Manager Items...\nLoading AECC Survey Manager Items...\nLoading AECC Plan Production Manager Items...\nLoading AECC Building Site Manager Items..."))
  
  ;; --- PHẦN BỔ SUNG THEO YÊU CẦU ---
  
  ;; 1. Xóa lịch sử dòng lệnh bằng cách in nhiều dòng trống
  (repeat 50 (princ "\n"))
  
  ;; 2. Chạy lệnh SHOW_MENU
  (command "_.SHOW_MENU")
  
  ;; --- KẾT THÚC PHẦN BỔ SUNG ---
  
  ;; Kết thúc lệnh một cách "sạch sẽ" mà không hiển thị giá trị trả về
  (princ)
)