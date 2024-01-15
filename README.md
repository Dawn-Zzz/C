<div>
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

<!--<div class="container mt-5">
    <h2>Sửa Sản Phẩm</h2>

    @using (Html.BeginForm())
    {
        @Html.AntiForgeryToken()

        <div class="card">
            <div class="card-body">
                <h4 class="card-title">Sản Phẩm</h4>
                <hr />

                @Html.ValidationSummary(true, "", new { @class = "text-danger" })
                @Html.HiddenFor(model => model.MaSanPham)

                <div class="form-group row">
                    @Html.LabelFor(model => model.TenSanPham, htmlAttributes: new { @class = "col-md-2 col-form-label" })
                    <div class="col-md-10">
                        @Html.EditorFor(model => model.TenSanPham, new { htmlAttributes = new { @class = "form-control" } })
                        @Html.ValidationMessageFor(model => model.TenSanPham, "", new { @class = "text-danger" })
                    </div>
                </div>-->

                <!-- Các div.form-group và input/editor khác tương tự -->

                <!--<div class="form-group row">
                    <div class="col-md-offset-2 col-md-10">
                        <input type="submit" value="Lưu" class="btn btn-primary" />
                    </div>
                </div>
            </div>
        </div>
    }
    <div class="mt-3">
        @Html.ActionLink("Quay Lại", "Index", new { @class = "btn btn-secondary" })
    </div>
</div>-->
<!--<div class="container" style="margin-top:150px">
   <h2 class="mb-4">Chi Tiết Sản Phẩm</h2>

    <div class="card">
        <div class="card-body">
            <h4 class="card-title">Sản Phẩm</h4>
            <hr />

            <dl class="row">
                <dt class="col-sm-3">Tên Sản Phẩm</dt>
                <dd class="col-sm-9">@Html.DisplayFor(model => model.TenSanPham)</dd>

            </dl>
        </div>
    </div>

    <p class="mt-4">
        @Html.ActionLink("Sửa Sản Phẩm", "SuaSanPham", new { id = Model.MaSanPham }, new { @class = "btn btn-warning" }) |
        @Html.ActionLink("Quay Lại", "Index", new { @class = "btn btn-secondary" })
    </p>
</div>-->
</div>
