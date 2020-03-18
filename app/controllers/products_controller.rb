require 'axlsx'
class ProductsController < ApplicationController
  def index
    @products = Product.order('created_at DESC')
    p = Axlsx::Package.new
    wb = p.workbook
    wb.styles do |style|
      highlight_cell = style.add_style(bg_color: "EFC376",
                                       border: Axlsx::STYLE_THIN_BORDER,
                                       alignment: { horizontal: :center },
                                       num_fmt: 8)
      date_cell = style.add_style(format_code: "yyyy-mm-dd", border: Axlsx::STYLE_THIN_BORDER)
      wb.add_worksheet(name: "Products") do |sheet|
        @products.each do |product|
          sheet.add_row [product.title, product.price, product.created_at, product.updated_at], style: [nil, highlight_cell, date_cell]
        end
      end
    end
    respond_to do |format|
      format.xlsx do
        send_data p.to_stream.read, type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename: 'products.xlsx'
      end
    end
  end

  def show
    @product = Product.find(params[:id])
  end
end
