; Script copy file Civil3D2026.dll sang thư mục chỉ định
; Nếu có trùng tên thì sẽ copy đè

; Lấy đường dẫn thư mục chứa file script hiện tại
sourceDir := A_ScriptDir
sourceFile := sourceDir "\Civil3D2026.dll"

; Thư mục đích
destDir := "Y:\5.SOFT T27\1. FOR WORK\1. THIET KE DUONG\2.CIVIL 3D\2026\AutoCAD Civil 3D 2026 Win x64\x64\c3d"
destFile := destDir "\Civil3D2026.dll"

; Copy file với chế độ ghi đè (1 = overwrite nếu trùng tên)
FileCopy, %sourceFile%, %destFile%, 1

; Thông báo kết quả
if ErrorLevel
    MsgBox, 16, Lỗi, Không thể copy file %sourceFile% đến %destDir%
else
    MsgBox, 64, Thành công, Đã copy file Civil3D2026.dll đến thư mục đích thành công!
