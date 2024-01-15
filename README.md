{
[RegularExpression("^[a-zA-Z].*", ErrorMessage = "Ký tự đầu tiên phải là chữ cái.")]
[RegularExpression(@"^[0-9]+$", ErrorMessage = "Chỉ được nhập số")]
[RegularExpression(@"^.+\.jpg$", ErrorMessage = "Ảnh đại diện phải có đuôi là '.jpg'.")]
$(document).ready(function () {
    $('.arrivals_grid_sorting li').on('click', function () {
       
        // Lấy giá trị filter
        var filterValue = $(this).attr('data-filter');

        // Gửi Ajax request
        $.ajax({
            url: '@Url.Action("GetSanPhams", "Home")',
            type: 'GET',
            data: { key: filterValue },
            success: function (data) {
                // Cập nhật danh sách sản phẩm
                $('#listSanPham').html(data);
                initIsotope();
                console.log(data);
            },
            error: function (error) {
                console.log(error);
            }
        });
    });
});
}
