---
name: lam-slide
description: Hỗ trợ lên dàn ý và soạn nội dung slide thuyết trình bằng tiếng Việt — gợi ý cấu trúc deck, viết tiêu đề từng slide, tóm tắt ý chính thành bullet ngắn, đề xuất hình minh họa, giữ đúng giọng điệu cho người xem là giáo viên. Mặc định dành cho chủ đề kỹ thuật cơ khí, đặc biệt là hệ thống gá kẹp (đồ gá, ê tô, cơ cấu kẹp, định vị) trên máy phay CNC. Hãy dùng skill này mỗi khi người dùng nói "tạo slide", "làm slide", "lên dàn ý bài trình bày", "chuẩn bị deck", "soạn bài thuyết trình", "slide về gá kẹp / đồ gá / CNC", hoặc nhờ tóm tắt tài liệu thành slide, kể cả khi họ không nói rõ chữ "skill" hay "dàn ý".
---

# Làm slide cho bài trình bày kỹ thuật

Skill này giúp người dùng đi từ một chủ đề (thường là hệ thống gá kẹp trên máy phay CNC) đến một bộ khung slide hoàn chỉnh: dàn ý, tiêu đề, bullet, gợi ý hình. Người xem mặc định là **giáo viên**, phong cách **ngắn gọn, dễ hiểu**.

Skill chỉ soạn nội dung dạng văn bản. Nếu người dùng muốn xuất file .pptx, hãy làm xong nội dung theo skill này rồi dùng công cụ/skill tạo pptx nếu có.

## Bước 1: Nắm bối cảnh, ít hỏi nhất có thể

Cần biết: chủ đề cụ thể, thời lượng hoặc số slide, mục tiêu bài (giới thiệu, giảng giải, báo cáo, bảo vệ). Nếu người dùng chưa nói, **tự giả định rồi nêu rõ giả định ở đầu câu trả lời** thay vì hỏi dồn một loạt câu. Người dùng chỉnh lại sau còn nhanh hơn là bắt họ trả lời trước. Chỉ hỏi lại khi chủ đề quá mơ hồ để bắt đầu (ví dụ chỉ nói "làm slide" mà không có chủ đề nào).

Giả định mặc định nếu thiếu thông tin:
- Người xem: giáo viên (hiểu kỹ thuật cơ bản, quen với cách trình bày có hệ thống, sẽ để ý độ chính xác của thuật ngữ và logic sư phạm)
- Thời lượng: 10–15 phút, khoảng 8–12 slide (xấp xỉ 1 slide/1–1,5 phút)
- Ngôn ngữ: tiếng Việt, kèm thuật ngữ tiếng Anh trong ngoặc ở lần đầu xuất hiện, ví dụ "ê tô (vise)"

## Bước 2: Dàn ý tổng

Đưa dàn ý trước, mỗi dòng một slide, để người dùng duyệt luồng bài trước khi đi vào chi tiết. Lý do: sửa thứ tự slide ở giai đoạn này rẻ hơn nhiều so với sửa sau khi đã viết xong bullet.

Khung mặc định, co giãn theo thời lượng:
1. Tiêu đề + mục tiêu bài (người xem sẽ hiểu được gì sau bài này)
2. Vấn đề / vì sao chủ đề quan trọng
3. Khái niệm nền (định nghĩa, thuật ngữ)
4. Nội dung chính, chia 3–5 slide, mỗi slide một ý
5. So sánh hoặc ví dụ ứng dụng
6. Lưu ý, lỗi thường gặp, an toàn
7. Tóm tắt (3 ý cần nhớ)
8. Câu hỏi thảo luận / Hỏi đáp, tài liệu tham khảo

Với bài về gá kẹp CNC, đọc `references/gia-kep-cnc.md` để lấy khung nội dung kỹ thuật (nguyên lý định vị, các loại cơ cấu kẹp, tiêu chí chọn, lỗi thường gặp). Đây là gợi ý để chọn lọc, không phải danh sách phải nhồi hết vào slide.

## Bước 3: Soạn từng slide

Mỗi slide trình bày theo đúng khuôn này:

```
### Slide N — [Tiêu đề]
- [bullet 1]
- [bullet 2]
- [bullet 3]
**Hình minh họa:** [mô tả cụ thể]
**Ghi chú nói:** [1–2 câu, chỉ khi cần]
```

Nguyên tắc viết, kèm lý do:

- **Một slide một ý.** Người xem vừa nghe vừa đọc; nhiều ý trên một slide làm họ không theo kịp.
- **Tiêu đề là nhận định, không phải nhãn.** "Ê tô kẹp nhanh nhưng giới hạn hình dạng phôi" thông tin hơn "Ê tô". Giáo viên đọc tiêu đề là nắm được thông điệp.
- **3–5 bullet, mỗi bullet tối đa khoảng 12 từ**, bắt đầu bằng từ khóa chính. Cụm danh từ hoặc câu rất ngắn, không viết đoạn văn. Chi tiết dài chuyển sang "Ghi chú nói".
- **Dùng thuật ngữ chuẩn, nhất quán.** Một khái niệm chỉ gọi một tên xuyên suốt deck (đã chọn "đồ gá" thì không đổi sang "đồ kẹp" giữa chừng). Giáo viên dễ phát hiện và bị phân tâm.
- **Định nghĩa trước, dùng sau.** Thuật ngữ mới phải được giải thích ở slide đầu tiên dùng đến nó.
- **Có logic dẫn dắt.** Slide sau nối tiếp slide trước (vấn đề → nguyên lý → giải pháp → so sánh → lưu ý). Nếu có thể, tiêu đề đọc liền một mạch phải kể được cả câu chuyện của bài.
- **Số liệu, thông số chỉ đưa khi chắc chắn.** Với lực kẹp, dung sai, vật liệu… nếu không chắc con số thì dùng mô tả định tính hoặc ghi rõ "cần kiểm tra theo tài liệu/nhà sản xuất". Đừng bịa số; người xem là giáo viên sẽ hỏi nguồn.
- **Giọng điệu:** rõ ràng, khách quan, vừa phải trang trọng; không sáo rỗng, không khoa trương. Giải thích đơn giản nhưng không làm đơn giản đến mức sai.

## Bước 4: Gợi ý hình minh họa

Với mỗi slide, đề xuất **một** hình phục vụ đúng ý của slide đó, mô tả đủ cụ thể để người dùng tìm hoặc vẽ được. Nêu rõ loại hình:

- **Sơ đồ tự vẽ** (nên chọn khi là khái niệm trừu tượng): ví dụ "hình phôi hình hộp với 3 mặt định vị: 3 điểm tựa mặt đáy, 2 điểm mặt bên, 1 điểm mặt đầu, đánh số và mũi tên lực kẹp"
- **Ảnh thực tế**: ví dụ "ảnh ê tô máy phay kẹp phôi nhôm, nhìn chéo từ trên xuống"; kèm **từ khóa tìm kiếm** (cả tiếng Việt và tiếng Anh, như "CNC milling vise workholding", "3-2-1 locating principle fixture")
- **Ảnh chụp CAD/mô phỏng**: nếu người dùng có mô hình đồ gá
- **Bảng so sánh**: khi slide đặt nhiều phương án cạnh nhau, bảng thường rõ hơn hình ảnh

Không mô tả hình chung chung kiểu "ảnh minh họa đồ gá"; hình phải trả lời được "nó giúp người xem hiểu điều gì?". Không bịa đường dẫn hoặc nguồn ảnh cụ thể; chỉ gợi ý loại hình và từ khóa tìm, hoặc nhắc người dùng kiểm tra bản quyền ảnh khi lấy từ internet. Slide chỉ toàn chữ mà không cần hình thì ghi "không cần hình".

## Bước 5: Kiểm tra trước khi giao

Rà lại nhanh, sửa trước khi trả lời:
- Đọc riêng dãy tiêu đề: có kể thành một mạch không?
- Có slide nào quá 5 bullet hoặc bullet quá dài không?
- Thuật ngữ có nhất quán và được định nghĩa trước khi dùng không?
- Có con số nào mình không chắc không? Nếu có, đổi sang diễn đạt định tính hoặc đánh dấu cần kiểm chứng.
- Tổng số slide có khớp thời lượng đã giả định không?

Kết thúc bằng một câu mời người dùng chỉnh: đổi độ sâu, thêm/bớt slide, hoặc xuất sang file pptx.

## Khi người dùng đưa tài liệu sẵn

Nếu họ dán hoặc đưa file giáo trình/bài báo/ghi chú: rút ý chính từ đó, giữ nguyên thuật ngữ trong tài liệu gốc, và không thêm thông tin kỹ thuật ngoài tài liệu mà không nói rõ là phần bổ sung.
