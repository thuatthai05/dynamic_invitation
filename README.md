# dynamic_invitation

 

Tên dự án : TỰ ĐỘNG HÓA VIỆC TẠO THƯ MỜI VÀ GỮI EMAIL TÙY CHĨNH  VỚI WORDS-EXCEL VÀ PYTHON
Tác giả: Thuat Thai
Email : thuat.thai@thaihothuat.name.vn
Website: thaihothuat.name.vn
Dự án nhỏ này được tạo ra nhằm giúp các bạn làm văn phòng, người làm kinh doanh, sale, marketing… thực hiện các hoạt động gửi thư mời, quảng bá sản phẩm, dịch vụ … của mình một cách tự động, hàng loạt ,nhanh chóng và chính xác. Đây là một dự án mã nguồn mở, miễn phí , các bạn có thể sữ dụng cho bất kỳ mục đích nào mà không cần hỏi ý kiến tác giả.
Tuy nhiên, tác giả không khuyến khích và cũng không chịu trách nhiệm nếu các bạn sữ dụng tài nguyên cũng như công cụ của dự án vào các mục đích gây ảnh hưởng xấu đến người khác và cộng đồng như spam email, gửi thông tin lừa gạt hàng loạt…
Tác giả hy vọng mọi người sử dụng sản phẩm của dự án với mục đích tốt đẹp cho cuộc sống, mang lại giá trị cho cá nhân người sữ dụng, không gây nguy hại cho cộng đồng, đồng thời đóng góp sự hiểu biết vào sự phát sự phát triển của toàn xã hội.
Các ứng dụng chính của dự án:
	Tạo file words hàng loạt (như thư mời, thiệp cưới, danh thiếp….) tùy chĩnh  (với dữ liệu lưu trong file excel) từ file words  mẫu.
	Xuất sang file *.pdf  hàng loạt.
	Gửi email hàng loạt đến một danh sách email , đính kèm thư mời words.
	Gửi email hàng loạt đến một danh sách email , đính kèm thư mời *pdf.

I.Chuẩn bị tài nguyên và công cụ cho dự án:
1.Đầu tiên bạn cần download toàn bộ source code của dự án về máy tính của bạn và giải nén vào 1 thư mục trên máy tính của bạn.
Toàn bộ  mã nguồn - source code của dự án được public tại liên kết sau:
https://github.com/ThuatThai/create_words-file_and_send_email_automatically/tree/main
 

2.Ngoài Ms-words và MS- Excel,nếu máy tính của bạn chưa cài đặt python hãy tiến hành tải và cài đặt nó theo liên kết:
https://www.python.org/downloads/
3.Sau khi cài đặt python, bạn cần cài thêm một số công cụ sau:
PIP :  Vào thư mục sau khi tải về và giải nén, chọn chuột phải và chọn “open in terminal” , một khung cmd màu đen xuất hiện như sau:
 
Nhập vào dòng nhắt lệnh dòng lệnh sau để cài đặt: pip install pip
Python-docx: Nhập vào dòng nhắt lệnh dòng lệnh sau để cài đặt: pip install python-docx
(Đây là thư viện cho phép python hoạt động với words)
Openpyxl: Nhập vào dòng nhắt lệnh dòng lệnh sau để cài đặt: pip install openpyxl
(Đây là thư viện cho phép python hoạt động với excel)
smtplib: Nhập vào dòng nhắt lệnh dòng lệnh sau để cài đặt: pip install smtplib
(Đây là thư viện cho phép python hoạt động với EMAIL)
email: Nhập vào dòng nhắt lệnh dòng lệnh sau để cài đặt: pip install email
(Đây là thư viện cho phép python hoạt động với EMAIL)

os: Nhập vào dòng nhắt lệnh dòng lệnh sau để cài đặt: pip install os
(Đây là thư viện cho phép python hoạt động với thư mục và tệp tin)

Qúa trình chuẩn bị công cụ đã xong , ta tiến hành sữ dụng tài nguyên của dự án.
II.Cách sữ dụng:
Để sử dụng ứng dụng này cho dự án của bạn, bạn cần thay đổi dữ liệu trong tệp tin "template.docx" và "data.xlsx" thành dữ liệu của riêng bạn. Dưới đây là quy trình chi tiết để chuẩn bị dữ liệu cho dự án:
1.	Tệp tin Word mẫu (template.docx):
o	Mở tệp tin "template.docx" bằng một trình soạn thảo Word như Microsoft Word hoặc Google Docs.
o	Thay đổi nội dung email trong tệp tin Word để phù hợp với nhu cầu của bạn. Bạn có thể thay đổi văn bản, định dạng, thêm hình ảnh hoặc các thành phần khác cần thiết.
o	Lưu lại các thay đổi và đóng tệp tin Word.
Chú ý : 
	Bạn có thể thay tệp này bằng bất kỳ mẫu thư mời nào bạn muốn, tuy nhiên sau đó hãy ghi đè file lại với tên template.docx .
	Bạn có thể thay đổi hoặc thêm, bớt vị trí của các thông tin tùy chĩnh được đánh dấu màu đỏ trong file words , tuy nhiên nếu bạn thay đổi hoặc thêm vào các thông tin này, chú ý hãy cập nhật nó trong file data.xlsx

2.	Tệp tin Excel chứa dữ liệu (data.xlsx):
o	Mở tệp tin "data.xlsx" bằng một trình chỉnh sửa Excel như Microsoft Excel hoặc Google Sheets.
o	Trong tệp tin Excel, bạn sẽ thấy một bảng dữ liệu với các cột và hàng. Hàng đầu tiên  là tiêu đề cột, trong khi các hàng tiếp theo chứa dữ liệu của từng người nhận.
o	Thay đổi dữ liệu trong các cột tương ứng với thông tin bạn muốn sử dụng trong email, bao gồm địa chỉ email người nhận. Đảm bảo rằng mỗi hàng chứa thông tin của một người nhận và địa chỉ email phải nằm trong cột riêng.HÃY SỮ DỤNG ĐỊA CHỈ EMAIL CHÍNH XÁC VÌ CODE SẼ KHÔNG CHẠY NẾU EMAIL KHÔNG CHÍNH XÁC
o	Lưu lại các thay đổi và đóng tệp tin Excel.
Chú ý: Cập nhật lại tiêu đề cột ở hàng đầu sao cho trung khớp với thông tin trong file words
 
3.	 Kiểm tra lại việc cài đặt thư viện Python cần thiết: Đảm bảo bạn đã cài đặt các thư viện Python cần thiết như pip,openpyxl,docx, smtplib, email, và os. Bạn có thể cài đặt các thư viện này bằng cách sử dụng pip, ví dụ: pip install docx.

4.	Thiết lập thông tin tài khoản email gửi: Để gửi email, bạn cần cung cấp thông tin tài khoản email người gửi, bao gồm địa chỉ email và mật khẩu. Hãy chắc chắn rằng bạn có thông tin đăng nhập cho tài khoản email người gửi.

5.	 Thư mục chứa tệp tin: Một thư mục trên máy tính của bạn và di chuyển các tệp tin Word đã chuẩn bị vào thư mục đó.

6.	  Cấu hình và chạy mã: Mở trình biên dịch Python hoặc môi trường phát triển và chạy mã Python đã được cung cấp ở trên. Đảm bảo bạn đã cập nhật các biến sender_email, sender_password và directory cho phù hợp với thông tin và đường dẫn của bạn.

Các file mã Python (file có đuôi .py) được viết có các ứng dụng sau:
1.	Mã Python để tạo file Word (dynamic_invitation.py):
o	Ứng dụng: Khi chạy file này cho phép tạo các tệp tin Word với nội dung tùy chỉnh.
o	Lợi ích: Người dùng có thể tạo ra các tài liệu Word dễ dàng và tự động từ file mẫu template.docx
2.	Mã Python để chuyển đổi file Word sang PDF và xóa file Word (convert_files_to_pdf_and_delete_words_files.py):
o	Ứng dụng: Khi chạy file này cho phép chuyển đổi các tệp tin Word thành định dạng PDF và sau đó xóa các tệp tin Word gốc.
o	Lợi ích: Chuyển đổi sang định dạng PDF giúp đảm bảo tính nhất quán và tiện lợi cho việc chia sẻ và xem tài liệu. Xóa các tệp tin Word gốc giúp giải phóng không gian lưu trữ.
3.	Mã Python để gửi email hàng loạt đính kèm file Word(send_email_automatically_attach_words.py):
o	Ứng dụng Khi chạy file này cho phép gửi email hàng loạt với nội dung từ các tệp tin Word được tạo ra.
o	Nếu bạn đã xóa các file words ở bước 2 , chỉ có thể chạy mã python ở bước 4.Bước này sẽ gây lỗi. Do đó nếu muốn gửi file words thì không chạy bước 2.
o	Lợi ích: Người dùng có thể gửi email cá nhân hóa hoặc thông báo đến một nhóm người nhận với tệp tin Word đi kèm. Điều này hữu ích trong việc gửi thư mời, thông báo, hoặc nội dung tùy chỉnh cho nhiều người dùng.
4.	Mã Python để gửi email hàng loạt đính kèm file PDF (send_email_automatically_attach_pdf.py):
o	Ứng dụng: Khi chạy file này cho phép gửi email hàng loạt với nội dung từ các tệp tin PDF đã được chuyển đổi.
o	Lợi ích: Người dùng có thể gửi email cá nhân hóa hoặc thông báo đến một nhóm người nhận với tệp tin PDF đi kèm. Điều này hữu ích trong việc chia sẻ tài liệu, biểu mẫu hoặc báo cáo theo định dạng PDF.
Cách chung để chạy các file python (*.py):
Tại khung cmd gõ lệnh : python ten_file.py .
Ví dụ : Muốn chạy file dynamic_invitation.py:
 
Chú ý: 
Để cho phép tài khoản Gmail tự động đăng nhập từ mã Python, bạn cần thực hiện các bước sau:
1.	Cho phép truy cập ứng dụng kém an toàn:
o	Truy cập vào tài khoản Gmail của bạn trên trình duyệt web.
o	Vào phần Cài đặt (Settings) của tài khoản Google.
o	Chọn mục Bảo mật (Security).
o	Cuộn xuống và tìm mục "Truy cập ứng dụng kém an toàn".
o	Bật tùy chọn "Cho phép truy cập ứng dụng kém an toàn".
2.	Vào phần “ứng dụng khác” tạo một ứng dụng mới. Gmail sẽ tự động tạo cho bạn một mật khẩu thay thế mật khẩu của bạn.
3.	Sữ dụng mật khẩu này và điền vào thay thế mật khẩu của bạn trong tệp “send_email_automatically_attach_words.py” và tệp “send_email_automatically_attach_words.py” 
(Bạn có thể sữ dụng một trình đọc file text như notepad để đọc các file đuôi .py)
 

Trên đây là toàn bộ cách sữ dụng dự án “Ứng dụng python trong tự động hóa việc tạo thư mời và gửi email hàng loạt”.
Hy vọng dự án sẽ giúp các bạn giảm 1 phần thời gian, công sức cũng như tăng tính chính xác khi tạo thư mời, thiệp…với words và excel, tăng năng suất công việc .
Một lần nữa tác giả hy vọng các bạn sữ dụng đúng mục đích, không sữ dụng các công cụ được cung cấp trong dự án để spam email, gây phiền toái hay lừa gạt cộng đồng , xã hội.
Chân thành cảm ơn sự quan tâm của các bạn đến dự án.
Thân chào,
Thuật,
 


